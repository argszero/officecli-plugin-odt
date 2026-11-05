#!/usr/bin/env python3
"""Fetch the corpus documents described in sources.json, pinned and verified.

The documents are **not** vendored in this repository. They come from projects
whose licences (GPL-2.0, GPL-2.0-or-later, MPL-2.0) do not permit dropping their
bytes into an MIT repository, so what is committed is the manifest: repository,
commit, path and sha256. This script reconstitutes the exact bytes, and refuses
anything that does not hash to the recorded value.

    python3 tests/corpus/fetch.py            # fetch what is missing
    python3 tests/corpus/fetch.py --force    # re-fetch everything
    python3 tests/corpus/fetch.py --list     # show what the manifest names

Two transports, because the machines differ: `gh api` where the GitHub CLI is
installed and authenticated (works behind this project's sandbox), and a plain
HTTPS GET of raw.githubusercontent.com at the pinned commit otherwise (the CI
runner). Both are pinned to the commit, so a moving branch cannot change what a
run measures.
"""
import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
FILES = os.path.join(HERE, "files")
SOURCES = os.path.join(HERE, "sources.json")


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def documents():
    with open(SOURCES, encoding="utf-8") as fh:
        return json.load(fh)["documents"]


def via_gh(doc, dest):
    url = "repos/%s/contents/%s?ref=%s" % (doc["repo"], doc["path"], doc["commit"])
    # Prompting is disabled and the call is bounded: on a CI runner `gh` is
    # installed but unauthenticated, and a transport that waits for a prompt
    # would hang the build instead of falling through to HTTPS.
    env = dict(os.environ, GH_PROMPT_DISABLED="1", GH_NO_UPDATE_NOTIFIER="1")
    try:
        p = subprocess.run(["gh", "api", "-H", "Accept: application/vnd.github.raw", url],
                           capture_output=True, timeout=120, env=env)
    except (subprocess.TimeoutExpired, OSError) as exc:
        return False, "%s: %s" % (type(exc).__name__, exc)
    if p.returncode != 0 or not p.stdout:
        return False, (p.stderr or b"").decode("utf-8", "replace").strip()[:200]
    with open(dest, "wb") as fh:
        fh.write(p.stdout)
    return True, ""


def via_https(doc, dest):
    url = "https://raw.githubusercontent.com/%s/%s/%s" % (
        doc["repo"], doc["commit"], urllib.parse.quote(doc["path"]))
    try:
        with urllib.request.urlopen(url, timeout=60) as resp:
            data = resp.read()
    except Exception as exc:                                  # noqa: BLE001
        return False, "%s: %s" % (type(exc).__name__, exc)
    with open(dest, "wb") as fh:
        fh.write(data)
    return True, ""


def fetch(doc, force=False):
    dest = os.path.join(FILES, doc["file"])
    if os.path.exists(dest) and not force:
        if sha256(dest) == doc["sha256"]:
            return "kept", ""
        os.unlink(dest)
    transports = []
    if shutil.which("gh"):
        transports.append(via_gh)
    transports.append(via_https)
    errors = []
    for fn in transports:
        ok, err = fn(doc, dest)
        if not ok:
            errors.append("%s: %s" % (fn.__name__, err))
            continue
        got = sha256(dest)
        if got != doc["sha256"]:
            os.unlink(dest)
            errors.append("%s: sha256 %s != %s" % (fn.__name__, got[:16], doc["sha256"][:16]))
            continue
        return "fetched", ""
    return "failed", " | ".join(errors)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--list", action="store_true")
    a = ap.parse_args(argv)
    docs = documents()
    if a.list:
        for d in docs:
            print("%-28s %8d B  %-22s %s" % (d["file"], d["bytes"], d["producer"], d["license"]))
        print("\n%d documents from %d producers"
              % (len(docs), len({d["producer"] for d in docs})))
        return 0
    os.makedirs(FILES, exist_ok=True)
    counts = {"kept": 0, "fetched": 0, "failed": 0}
    for d in docs:
        state, err = fetch(d, a.force)
        counts[state] += 1
        if state == "failed":
            print("FAILED  %s\n        %s" % (d["file"], err))
        elif state == "fetched":
            print("fetched %s" % d["file"])
    print("\n%d kept, %d fetched, %d failed" % (counts["kept"], counts["fetched"],
                                                counts["failed"]))
    return 1 if counts["failed"] else 0


if __name__ == "__main__":
    sys.exit(main())
