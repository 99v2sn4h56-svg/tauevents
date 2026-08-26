# TAU Events — Apps Script

Apps Script project bound to the "Event Request Form - The Arts Unit (Responses)" Google Sheet.

## Setup

```bash
npm install
npx clasp login   # one-time, opens browser for Google auth
```

## Workflow

- Edit `Code.js` (and any other script files) locally in VS Code.
- `npx clasp push` — push local changes up to the live Apps Script project.
- `npx clasp pull` — pull down changes made in the Apps Script web editor.
- `git add -A && git commit -m "..."` then `git push` — version control as usual.

`clasp push` and `git commit`/`push` are separate actions — pushing to Apps Script does not commit to git, and committing to git does not push to Apps Script. Do both when you make a change.

## Files

- `Code.js` — the script itself.
- `appsscript.json` — Apps Script project manifest (scopes, runtime version).
- `.clasp.json` — points this folder at the live Apps Script project (scriptId). Not a secret; auth tokens live outside this repo in `~/.clasprc.json`.
