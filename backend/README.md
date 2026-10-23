# Invite backend (Google Apps Script)

Bound to the **Engagement Party Guest List** Sheet. `Code.gs` is the whole backend.

## Deploy (once, about 3 minutes)

1. Open the Sheet, then **Extensions > Apps Script**.
2. Replace the contents of `Code.gs` with this folder's `Code.gs`. Save.
3. In `setAccessCode()`, change `'CHANGE-ME'` to the access code, select `setAccessCode` in the toolbar, and click **Run**. Approve the permissions prompt. Then change it back to `'CHANGE-ME'` so the code isn't stored in the script.
4. Select `syncSlugs` and click **Run**. A **Responses** tab appears with a readable slug per invite (`joseph-noura`). Edit any slug you like; links are `marchilino.com/<slug>`.
5. **Deploy > New deployment > Web app**. Execute as: *Me*. Who has access: *Anyone*. Copy the `/exec` URL.

After editing `Code.gs` later, use **Deploy > Manage deployments > Edit > New version** so the URL stays the same.

## Test guests

Run `addTestGuests` to add `alex-testguest` (solo) and `jordan-sam` (with plus-one). Run `removeTestGuests` to delete them from Invites and Responses (Guest List rows are not touched).

## Local test

```
node backend/test/run.mjs
```

Runs `Code.gs` against an in-memory copy of the Sheet layout: gate lookups, case-insensitive code, RSVP validation, lock after submit, close after 31 Dec 2026, and the writes to Responses, Invites and Guest List.
