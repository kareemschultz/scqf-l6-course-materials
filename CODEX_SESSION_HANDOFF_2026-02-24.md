# Codex Session Handoff (2026-02-24)

This file is a resume point for the next Codex session.

## Scope Covered This Session

1. Audited SCQF repo structure, assignment outputs, and rubric alignment.
2. Cross-checked current MasjidConnect app status (code, CI, deploy, live checks).
3. Captured concrete risks and next actions so work can continue immediately.

---

## Repo Status Snapshot

- SCQF repo: `/home/karetech/SCQF-L6-Course-Materials`
- Branch: `master`
- Working tree state at capture: clean
- Remote: `origin https://github.com/kareemschultz/SCQF-L6-Course-Materials.git`

---

## SCQF Audit Summary

### What exists

- Assignment folders and artifacts are present for:
  - `HE9E_46_Contemporary_Business`
  - `J229_76_Understanding_Business`
  - `J22A_76_Management_People_Finance`
  - `F1FE_12_Word_Processing_Presenting`
  - `F1FJ_12_Spreadsheet_Database`
- Final submission bundle exists:
  - `My_Assignments/Final_2026-02-02`

### Confirmed duplicates (business units)

`Final_2026-02-02` docs are byte-identical to unit `Final/` docs:
- J229 hash match
- J22A hash match
- HE9E hash match

### High-priority findings

1. IT submissions contain many literal screenshot placeholders.
   - Affected files:
     - `My_Assignments/F1FE_12_Word_Processing_Presenting/Final/252IFCBR0596_KareemSchultz_F1FE_12_Word_Processing_Presenting.docx`
     - `My_Assignments/F1FJ_12_Spreadsheet_Database/Final/252IFCBR0596_KareemSchultz_F1FJ_12_Spreadsheet_Database.docx`
     - `My_Assignments/F1FJ_12_Spreadsheet_Database/Final/Hospital_Database_Documentation.docx`
     - `My_Assignments/F1FJ_12_Spreadsheet_Database/Final/Library_Database_Documentation.docx`
   - Evidence snapshot:
     - F1FE report: `placeholders:19`, `media:1`
     - F1FJ report: `placeholders:19`, `media:1`
     - Hospital DB doc: `placeholders:5`, `media:0`
     - Library DB doc: `placeholders:5`, `media:0`

2. Macro deliverable risk in F1FJ.
   - Excel files are `.xlsx` without VBA payload (`vbaProject.bin` not found).
   - If assessor expects a recorded macro file, this can lose marks.
   - Files checked:
     - `Employee_Training.xlsx`
     - `Sales_Expenses_Analysis.xlsx`
     - `StoreData_Analysis.xlsx`

3. F1FJ wording drift in report Task B (database section).
   - Main report text references `Treatments` and `Doctors` entities in sections, while official brief centers on `Patients` + `TreatmentRecords` for required query output.
   - A strict marker may penalize mismatch in evidence narrative.

### Medium-priority findings

4. J22A final still contains unresolved note:
   - `[NOTE: For the final .docx submission, replace this with a proper pyramid diagram image.]`
   - Present in:
     - `My_Assignments/J22A_76_Management_People_Finance/Final/J22A_76_Final_Assignment.docx`
     - `My_Assignments/Final_2026-02-02/252IFCBR0596_KareemSchultz_J22A_76_Management_People_Finance.docx`

5. Cover/declaration fields appear unfinalized in several docs.
   - Tutor/date/signature placeholders remain in extraction checks.
   - If those were final upload files, submission risk remains.

### Low-priority findings

6. Docs metadata drift:
   - `CLAUDE.md` and assignment readmes still emphasize old deadline framing.
   - Top-level `README.md` "Latest final submissions" only lists 3 business docs and omits IT package context.

7. Script portability issue:
   - `Scripts/prepare_final_submissions.py` uses hardcoded Windows path and inline pip install behavior.

---

## SCQF: Access / Database Capability Notes

Environment currently has no native MS Access toolchain installed (`mdbtools`, Access runtime, LibreOffice DB tooling not detected).

What can still be done immediately:
- rubric-aligned DB design review,
- SQL/query correction,
- CSV integrity and relationship validation,
- documentation cleanup and evidence mapping.

What needs setup for direct `.accdb` object inspection/editing:
- install `mdbtools` (or Windows Access environment / compatible runtime).

---

## MasjidConnect App Context (Cross-Project)

### Repo

- Path: `/home/karetech/v0-masjid-connect-gy`
- Branch: `main`
- HEAD at capture: `b623694` (`Expand remaining truncated belief references`)
- Working tree at capture:
  - modified: `next-env.d.ts`
  - untracked: `tmp/`

### Recent app commits relevant to this session chain

- `b623694` Expand remaining truncated belief references
- `c267660` Replace placeholder ellipses in Islamic content text
- `f2582ed` Set DISABLE_RATE_LIMIT in spawned test app
- `1f393be` Disable API rate limiting during test runs
- `1a1560f` Fix CI lint dependency and stabilize test suite
- `e51e0ab` Revamp adhkar/hadith/fiqh and add tafseer study hub
- `f1798cf` Add map-specific hero motion and premium loading states
- `6c73129` Unify premium UI across masjid and iftaar pages

### CI status

Latest run checks showed green:
- `22368441647` success (CI on `main`, push for `b623694`)
- Prior two runs also success after fixes.

### Deploy/live status at capture

- Container:
  - `kt-masjidconnect-prod`
  - image: `ghcr.io/kareemschultz/v0-masjid-connect-gy:latest`
  - status: running
- Live endpoint checks confirmed updated belief quote text is present:
  - `/explore/new-to-islam/beliefs` shows full expanded verses (no old truncated `...` variant).

### Open app backlog from user requests (still pending unless separately completed)

1. Tour guide performance/positioning issues on phone (hang/delayed/out-of-place steps) with Playwright real mobile simulation.
2. More custom, page-specific premium animations; remove generic/reused motion in certain hero/sections.
3. Buddy system UX: clear error when trying to add non-existent user.
4. Continue premium UI consistency polish across masjid, masjid detail, iftaar surfaces and other out-of-place screens.
5. Expand Dua/Adhkar/Hadith/Tafseer/Fiqh depth and completeness from GII library content sources.
6. Keep CI monitoring stable (email noise was from earlier failed runs; current recent runs are passing).

---

## Recommended Next Session Start Order

1. SCQF remediation pass (high impact, fast):
   - remove screenshot placeholders in IT docs and insert real evidence,
   - resolve J22A `[NOTE]`,
   - verify cover/declaration completion state.
2. F1FJ macro/evidence hardening:
   - provide an `.xlsm` with actual macro module (or explicit assessor-accepted workaround evidence).
3. MasjidConnect backlog:
   - start with tour guide mobile reliability using Playwright scripted reproduction and timing diagnostics.

---

## Quick Resume Commands

```bash
# SCQF repo
cd /home/karetech/SCQF-L6-Course-Materials
git status --short --branch

# Masjid app repo
cd /home/karetech/v0-masjid-connect-gy
git status --short --branch
gh run list --limit 8
docker ps --filter name=kt-masjidconnect-prod
```

