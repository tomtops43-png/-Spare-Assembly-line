# CLAUDE.md

H9 Spare Parts Portal: a single-file SPA (`index.html`) served from GitHub Pages, with a Google Apps Script backend (`src/Backend.gs`).

## Layout
- `index.html`: the whole frontend (HTML + CSS + JS). Tailwind is the v2.2.19 static build, so `slate-*` colors and arbitrary values like `min-w-[700px]` don't exist. Use plain CSS for those.
- `src/*.gs`: backend code. `clasp push` uploads it from here.
- `appsscript.json` (repo root): the only manifest to edit. The deploy workflow copies it to `src/` (which is gitignored) before pushing. Keep `webapp.access = ANYONE_ANONYMOUS` and `executeAs = USER_DEPLOYING`, because the frontend calls the backend without a login.
- `.clasp.json`: its `scriptId` is a placeholder (`YOUR_ACTUAL_SCRIPT_ID`). CI supplies the real one from a secret. Never commit a real ID or `.clasprc.json`.
- `tests/`: plain-node tests, run with `npm test`. Run them as `TZ=Asia/Bangkok npm test` to match CI, because time-based tests are off by 7 hours in UTC.

## Deploy
- Merging to `main` deploys the frontend via `.github/workflows/deploy-pages.yml`.
- A merge to `main` that touches `src/**` or `appsscript.json` runs `.github/workflows/deploy-appsscript.yml`, which does `clasp push --force` and then `clasp deploy --deploymentId` so the `/exec` URL stays the same. Setup steps are in `DEPLOYMENT.md`.

## Pull requests
- **Auto-merge rule:** for a PR in this repo, once every CI check on the latest head commit has passed and no review is outstanding (no unresolved review threads and no "changes requested" review), squash-merge it yourself without asking. If it is still a draft, mark it ready for review first. Don't merge while any check is failing or still running, or while there is a merge conflict.
- The squash commit title follows the repo's existing style: `type: emoji short description (#PR)`.
