# I-O Psych Digest — web version

A single scrollable page built from `io_psych_digest.xlsx`. Designed for reading on a phone.

- `build_site.py` — reads the spreadsheet, writes `index.html`
- `index.html` — the whole site. One self-contained file, no CDN, no external requests.

Because everything is inlined, the page works on GitHub Pages, from a local file, and offline
once loaded. Roughly 145 KB over the wire.

---

## One-time setup

### 1. Create the repo

On GitHub, create a new **public** repo (Pages needs public unless you have a paid plan).
Name it something like `io-psych-digest`. Don't add a README — this folder already has one.

Then, in Terminal on your Mac:

```bash
cd ~/Desktop/"Daily Read"/digest-site
git init -b main
git add -A
git commit -m "Initial digest site"
git remote add origin https://github.com/YOUR-USERNAME/io-psych-digest.git
git push -u origin main
```

### 2. Turn on GitHub Pages

Repo → **Settings** → **Pages** → under "Build and deployment", set Source to
**Deploy from a branch**, branch `main`, folder `/ (root)`. Save.

After a minute your page is live at:

```
https://YOUR-USERNAME.github.io/io-psych-digest/
```

Open that on your phone and add it to your home screen — it behaves like an app.

### 3. Publishing

**Publishing has to happen on your Mac, not from Claude's sandbox.** Two hard limits make
automated pushing from the sandbox impossible:

- The sandbox can create and modify files in this folder but **cannot delete them**. Git needs
  to create and remove lock files (`.git/index.lock`) constantly, so the first git command
  leaves a lock behind and every command after it fails.
- `api.github.com` is blocked from the sandbox, so the GitHub API isn't a way around it.

So the morning task rebuilds `index.html` locally, and you publish. Double-click:

```
publish.command
```

It clears any stale lock files, rebuilds the page from the current spreadsheet, commits, pulls
anything you changed on github.com, and pushes. The Terminal window shows what happened and
waits for a keypress before closing.

If macOS blocks it the first time ("cannot be opened because it is from an unidentified
developer"), right-click → Open → Open, or run `chmod +x publish.command` once.

### 4. Optional: publish automatically each morning

To skip the double-click, have macOS run the same script on a schedule. Save this as
`~/Library/LaunchAgents/com.daphnehou.dailyread.plist`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>com.daphnehou.dailyread</string>
  <key>ProgramArguments</key>
  <array>
    <string>/bin/bash</string>
    <string>/Users/daph/Desktop/Daily Read/digest-site/publish.command</string>
  </array>
  <key>StartCalendarInterval</key>
  <dict><key>Hour</key><integer>8</integer><key>Minute</key><integer>0</integer></dict>
  <key>StandardOutPath</key><string>/tmp/dailyread.log</string>
  <key>StandardErrorPath</key><string>/tmp/dailyread.log</string>
</dict>
</plist>
```

Then load it once:

```bash
launchctl load ~/Library/LaunchAgents/com.daphnehou.dailyread.plist
```

8:00am gives the 7:10am digest task time to finish. The script skips its "press any key" pause
when it isn't run from a Terminal, so it exits cleanly. Check `/tmp/dailyread.log` if a morning
seems to have been missed — and note the job only fires if your Mac is awake.

Git needs to already have your GitHub credentials for this to work unattended. Run
`publish.command` by hand once first; if the push succeeds without prompting, the scheduled run
will too.

---

## Rebuilding by hand

```bash
cd ~/Desktop/"Daily Read"/digest-site
python3 build_site.py           # rebuild index.html only
python3 build_site.py --push    # rebuild, commit, pull --rebase, push
```

Needs `openpyxl` (`pip3 install openpyxl`).

Close the spreadsheet in Excel first if you've been editing it — an open file is still readable,
but you want your latest edits saved before the page is generated from it.

---

## How themes work

The spreadsheet's "Topic Area" column is written fresh each run and is very specific — most
values appear only once, which makes them useless as a filter. So `build_site.py` rolls each
article up into one or more of ~16 broad themes (keyword matching on topic + title), and those
drive the theme dropdown. The original specific label still shows as the chip on each card.

To adjust the buckets, edit `THEME_RULES` near the top of `build_site.py` and rebuild. Any
article matching nothing lands in "Other", so that's the bucket to watch when you add new
subject areas.
