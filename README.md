
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

### 3. Let the daily task push for you

The morning task rebuilds the page automatically, but it runs in a sandbox that has no access
to your GitHub login. To let it push, give the repo a token it can use.

1. GitHub → Settings → Developer settings → **Personal access tokens** → **Fine-grained tokens**
2. Generate a token scoped to **only this one repo**, with **Repository permissions →
   Contents: Read and write**. Nothing else. Set an expiry you're comfortable with.
3. Store it in the remote URL:

```bash
cd ~/Desktop/"Daily Read"/digest-site
git remote set-url origin https://YOUR-TOKEN@github.com/YOUR-USERNAME/io-psych-digest.git
```

**Worth knowing:** this writes the token in plain text into `.git/config` on your Mac. That file
is never committed, so it won't end up on GitHub — but anyone with access to your laptop could
read it. That's the tradeoff for unattended pushing. Keep the token scoped to this single repo
so the blast radius stays small, and revoke it on GitHub if you ever stop using this setup.

If you'd rather not store a token, skip this step. The task will still rebuild `index.html`
every morning and just tell you it couldn't push; you then run `git push` yourself whenever
you like.

---

## Rebuilding by hand

```bash
cd ~/Desktop/"Daily Read"/digest-site
python3 build_site.py           # rebuild index.html only
python3 build_site.py --push    # rebuild, then commit and push
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
