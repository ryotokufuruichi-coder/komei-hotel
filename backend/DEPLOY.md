# Backend (GAS) deploy — clasp

The website (this repo) auto-deploys via GitHub Pages, but `backend/Code.gs`
runs on **Google Apps Script**, a separate service that a plain `git push`
cannot touch. Deploy it from this repo with Google's **clasp** CLI — local,
one command. The repo is the source of truth.

## One-time setup
1. Turn ON the Apps Script API: https://script.google.com/home/usersettings
2. `npm install`             # installs clasp locally (devDependency)
3. `npm run clasp:login`     # opens a browser — sign in as the project OWNER
4. Copy `.clasp.json.sample` → `.clasp.json` and paste your **Script ID**
   (Apps Script editor → ⚙ Project Settings → "Script ID"):
   ```json
   { "scriptId": "…", "rootDir": "backend" }
   ```
   (`.clasp.json` is gitignored; `.clasprc.json`, your OAuth token, must NEVER be committed.)

## ⚠️ First time only — reconcile the repo with the LIVE project
The live Apps Script project may hold edits that never made it back into this
repo. Never blindly overwrite it.
1. `npm run clasp:pull`      # downloads live Code.gs (OVERWRITES backend/Code.gs) + appsscript.json
2. `git diff backend/Code.gs`
   - Only *your* intended repo changes show as missing on live → repo is ahead → safe.
   - Live has changes the repo lacks → merge them into the repo first.
3. `git checkout -- backend/Code.gs`   # restore the repo's intended version
4. `git add backend/appsscript.json && git commit -m "chore: add clasp manifest"`

## Deploy — every time after committing a backend change
```bash
npm run deploy:backend
```
Runs `clasp push -f` (upload) then `clasp deploy -i <deploymentId>`, which
updates the **same** web-app URL (`/exec …fnOZR`). Confirm the id first with
`npm run clasp:deployments` and update it in `package.json` if it differs.

## Rules
- Always reuse the EXISTING deployment id so the `/exec` URL never changes
  (creating a *new* deployment mints a new URL and breaks the site).
- Commit `backend/Code.gs` to git first, then `npm run deploy:backend`.

## This project's specifics (verified 2026-08-03)
- **Script**: container-bound to the "Komei Hotel 予約管理" spreadsheet, so it does
  NOT appear in `clasp list`. Use the Script ID directly in `.clasp.json`.
- **Owner**: `ryotoku.furuichi@yoshinarcorp.com` (Workspace, yoshinarcorp.com domain).
  Editor: `ryotofuru838@gmail.com`.
- **Apps Script API** must be turned ON *per account* at
  https://script.google.com/home/usersettings — do it for whichever account clasp
  logs in as. Reads (pull/deployments) can succeed before writes propagate.
- **Deploy is domain-restricted**: only an account in the SAME domain as the owner
  (`@yoshinarcorp.com`) may promote a deployment. The gmail editor account can
  `clasp push` (sync code + create a version) but gets
  "Only users in the same domain as the script owner may deploy this script."
- **External-edit notification**: pushing as the gmail editor (`ryotofuru838`, outside
  the org) makes Google email the owner ("Someone outside your organization has edited
  an Apps Script project attached to your document."). Use the in-org OWNER `ryotoku`
  to avoid it.
- **Working process (preferred, verified @40 on 2026-08-03)**:
  Log in clasp as the OWNER `ryotoku.furuichi@yoshinarcorp.com` (Apps Script API ON) →
  `npm run deploy:backend` does push + deploy in ONE command. No editor-UI step and no
  external-edit notification.
  Fallback only if ryotoku's API is unavailable: `clasp push` as `ryotofuru838`, then the
  owner promotes via the editor UI (Deploy → Manage deployments → `…fnOZR` → New version).
- **Filenames**: after the first `clasp push`, the main server file is `Code`
  (the old editor file `コード` was renamed by the sync). `fix` is an empty stub kept
  only for parity.
- **Verify a deploy** (read path always works): 
  `GET <exec>?action=get_reservation&id=<id>&token=<token>` → the reservation JSON
  now includes a `guests` array when the passport-era code is live.
