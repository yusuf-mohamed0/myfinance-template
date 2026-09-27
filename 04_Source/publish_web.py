# -*- coding: utf-8 -*-
r"""
PUBLISH_WEB - rebuild the site in 06_Web and push it to GitHub Pages
(the public URL comes from 03_System/config.json via mfconfig.py).

Usage: python publish_web.py [plan_month YYYY-MM]   (default: current month)
Steps: build_web.py -> git add/commit -> git push -> wait Pages build.
ASCII-only console output.

If github_owner / github_repo are not set in config.json yet, this script
stops early and prints the one-time GitHub setup steps.
"""
import os, sys, subprocess, time, datetime as dt

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import mfconfig
from mfconfig import site_url

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEB = os.path.join(BASE, "06_Web")
MONTH = sys.argv[1] if len(sys.argv) > 1 else dt.date.today().strftime("%Y-%m")
REPO = mfconfig.repo_slug()
SITE = site_url()


def run(cmd, cwd=WEB, ok_codes=(0,)):
    p = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    return p.returncode, (p.stdout or "") + (p.stderr or "")


def main():
    if not REPO:
        print("SETUP NEEDED: github_owner / github_repo are empty in 03_System/config.json")
        print("Run init_project.py (or fill those two keys), then publish again.")
        sys.exit(2)

    # 1. rebuild site
    rc, out = run([sys.executable, os.path.join(BASE, "04_Source", "build_web.py"), MONTH])
    print(out.strip())
    if rc != 0:
        print("BUILD FAILED")
        sys.exit(1)

    if not os.path.isdir(os.path.join(WEB, ".git")):
        print("NO GIT REPO in 06_Web yet - run init_project.py to create your GitHub repo.")
        sys.exit(2)

    # 2. commit
    run(["git", "add", "-A"])
    msg = "update " + dt.datetime.now().strftime("%Y-%m-%d %H:%M") + " (plan " + MONTH + ")"
    rc, out = run(["git", "commit", "-m", msg])
    if "nothing to commit" in out:
        print("no changes to publish")
    elif rc != 0:
        print("COMMIT FAILED:\n" + out)
        sys.exit(1)
    else:
        print("committed: " + msg)

    # 3. push
    rc, out = run(["git", "push", "origin", "main"])
    if rc != 0:
        print("PUSH FAILED:\n" + out)
        sys.exit(1)
    print("pushed to origin/main")

    # 4. wait for Pages build
    print("waiting for GitHub Pages build...")
    for _ in range(30):
        rc, out = run(["gh", "api", "repos/" + REPO + "/pages", "--jq", ".status"])
        status = out.strip().splitlines()[-1] if out.strip() else "?"
        if status == "built":
            print("STATUS: built")
            print("LIVE: " + SITE)
            return
        time.sleep(6)
    print("STATUS: still " + status + " - check later: " + SITE)


if __name__ == "__main__":
    main()
