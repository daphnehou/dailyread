#!/bin/bash
# Rebuild the digest page from the spreadsheet and publish it to GitHub.
#
# Double-click this file in Finder, or point a launchd job at it.
# It runs on your Mac, where git has your credentials and can manage its own
# lock files — neither of which is true inside Claude's sandbox.

cd "$(dirname "$0")" || exit 1
echo "── I-O Psych Digest ──────────────────────────────"
echo

pause_if_interactive() { [ -t 0 ] && read -n1 -r -p "Press any key to close… "; echo; }

# Claude's sandbox can create and modify files here but cannot delete them,
# so git lock files it leaves behind have to be cleared from this side.
rm -f .git/index.lock .git/HEAD.lock .git/objects/maintenance.lock 2>/dev/null
find .git/objects -name 'tmp_obj_*' -delete 2>/dev/null
rm -f _probe.txt 2>/dev/null

if ! python3 -c "import openpyxl" 2>/dev/null; then
  echo "Missing the openpyxl library. Run this once, then try again:"
  echo "    pip3 install openpyxl"
  echo
  pause_if_interactive
  exit 1
fi

echo "Rebuilding page from spreadsheet…"
if ! python3 build_site.py; then
  echo
  echo "Build failed — the page was not changed."
  pause_if_interactive
  exit 1
fi
echo

git add -A
if git diff --staged --quiet; then
  echo "No changes to publish. The site is already current."
  echo
  pause_if_interactive
  exit 0
fi

git commit -q -m "Digest update $(date +%Y-%m-%d)" || {
  echo "Commit failed."; pause_if_interactive; exit 1; }

# Pick up anything changed on github.com directly — editing a file in the web
# UI, or letting GitHub add a CNAME when you set a custom domain — so the push
# below is not rejected as out of date.
git pull --rebase -q origin main 2>/dev/null || \
  echo "Note: could not rebase on origin/main. If the push fails, run 'git pull --rebase' yourself."

if git push -q origin main 2>&1; then
  echo "Published. The site updates in about a minute:"
  echo "    https://daphnehou.github.io/dailyread/"
else
  echo "Push failed. Usually this means git needs your GitHub sign-in."
  echo "Run 'git push' in Terminal from this folder once to sort it out."
fi

echo
pause_if_interactive
