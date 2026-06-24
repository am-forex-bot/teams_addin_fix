# Teams Meeting Add-in Fix (lasting, version-proof)

This makes the **Teams Meeting** button reappear in Outlook and keeps it working,
for **any user on any laptop**, **without admin rights** and **without slowing
down logon**.

---

## Why the add-in kept breaking

The Teams Meeting add-in lives in each user's profile, in a **version-numbered**
folder, e.g.:

```
C:\Users\<user>\AppData\Local\Microsoft\TeamsMeetingAddin\1.0.24313.1\x64\Microsoft.Teams.AddinLoader.dll
```

Outlook only shows the add-in if that loader DLL is **registered** for the user
and its `LoadBehavior` is set to `3` (load + enabled).

Every time Teams updates, it drops the add-in into a **new version folder**
(`1.0.24313.1` -> `1.0.25xxx.x` -> ...). The old fix reinstalled an MSI into a
**hard-coded version path**, so the moment Teams updated, the path was wrong and
the fix broke. It also needed an **admin/UAC prompt** and **closed Outlook**.

## How this version fixes it for good

`Fix-TeamsMeetingAddin.ps1` does **not** care about version numbers or installers:

1. Finds the **newest version folder actually present** on the machine
   (handles both `TeamsMeetingAddin` and `TeamsMeetingAdd-in` spellings).
2. **Checks first** whether the add-in is already correctly registered to that
   exact version with `LoadBehavior = 3`.
   - If yes -> it does **nothing** and exits in a fraction of a second.
     *(This is what keeps logon fast.)*
   - If no -> it silently re-registers the loader DLL for the current user
     (`regsvr32 /s /n /i:user`), forces `LoadBehavior = 3`, and tells Outlook
     not to auto-disable it.
3. It **never closes Outlook** and **always exits cleanly**, so it is safe to run
   at logon. Changes show up the next time Outlook is started.

No admin rights. No UAC. No MSI. No hard-coded versions.

> The fix re-registers **whatever version Teams has already installed**. If the
> add-in folder doesn't exist at all (Teams/the add-in was never installed for
> that user), there's nothing to register - that's a Teams install problem, not
> something this script can create from nothing.

---

## The files

| File | What it is |
|------|------------|
| **`Fix-TeamsMeetingAddin.ps1`** | The actual fix. Everything else just launches this. |
| **`Install-TeamsAddinFix.ps1`** | **Recommended rollout.** Run once per machine as admin; sets it to run hidden at every user logon. |
| **`Uninstall-TeamsAddinFix.ps1`** | Undoes the installer (removes the scheduled task + files). |
| **`FixTeamsAddin_Manual.bat`** | **"Fix my machine now"** button - double-click, runs visibly, forces a fix. |
| **`FixTeamsAddin_Logon.bat`** | Alternative logon launcher for GPO logon scripts (runs the fix hidden). |
| **`Run-Hidden.vbs`** | Helper used by the logon `.bat` to run with no window flash. |

---

## How to use it

### A) Fix one machine right now (no admin)
Double-click **`FixTeamsAddin_Manual.bat`**. It runs visibly, applies the fix,
and tells you to restart Outlook. That's it.

### B) Roll it out to everyone (recommended)
On each machine (or via Intune / SCCM / GPO / your RMM), run **once as admin**:

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\Install-TeamsAddinFix.ps1"
```

This copies the fix to `C:\ProgramData\TeamsAddinFix\` and creates a scheduled
task that runs it **hidden, at every user logon, in that user's own non-admin
context**. Because the check exits instantly when things are fine, it does not
slow logon down - it only does work when a fix is actually needed.

To test immediately without logging off, as the test user:
```powershell
Start-ScheduledTask -TaskName 'TeamsMeetingAddinFix'
```

### C) Prefer a Group Policy logon script instead of a scheduled task?
Put this whole folder on a share all users can read, and point a **User
Configuration -> Logon** script at **`FixTeamsAddin_Logon.bat`**. It self-locates
the `.ps1`/`.vbs` next to it, so no copying into `System32` is needed.

> Use **either** B **or** C, not both.

---

## How to check it worked

- In Outlook: **New Meeting** should show **Teams Meeting** on the ribbon
  (restart Outlook first if it was open).
- Logs (per user): `%LOCALAPPDATA%\Company\`
  - `TeamsAddinFix.status` - one line, the **last run** result
    (`OK`, `FIXED`, `NO-ADDIN`, or `ERROR`).
  - `TeamsAddinFix.log` - history of fixes/errors only (healthy runs are not
    logged, to keep it clean).

## Troubleshooting

| Symptom | What it means / what to do |
|---------|----------------------------|
| Status shows `NO-ADDIN` | The add-in isn't installed for that user. Make sure **new Teams** is installed and has run at least once. |
| Button still missing after `FIXED` | Fully close and reopen Outlook. If Outlook had hard-disabled it, run `FixTeamsAddin_Manual.bat` (the `-Force` run clears Outlook's disabled-add-in list). |
| Status shows `ERROR` | Open `TeamsAddinFix.log` for the message. |
| Want to force a re-register | `powershell -ExecutionPolicy Bypass -File Fix-TeamsMeetingAddin.ps1 -Force -ShowConsole` |
