# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-file, static HTML prototype (`index.html`) for a Thai-language faculty web form: "ระบบยื่นขอค่าตอบแทนพิเศษตามเกณฑ์มาตรฐานภาระงานวิชาการ" — a submission system for academic staff at the Faculty of Engineering, KMUTNB (King Mongkut's University of Technology North Bangkok) to claim special monthly compensation (ค่าตอบแทนพิเศษ) tied to publication/output quotas defined by university regulation.

There is no build system, package manager, bundler, or test suite — everything (markup, CSS, JS) lives inline in `index.html`. There is nothing to install or compile.

## Running it

Just open the file in a browser, or serve it locally if you need routing/CORS-safe behavior:

```
python3 -m http.server 8000
# then open http://localhost:8000/index.html
```

There is no build, lint, or test command — verify changes by opening the page and exercising the 3-step wizard in-browser.

## Source-of-truth documents

The two Thai-language PDFs in the repo root are the actual university regulations the eligibility logic implements:

- `ข้อบังคับมาตรฐานภาระงานทางวิชาการฯ (ฉบับที่ 2)` — the regulation defining minimum academic workload/output standards per position (referenced in code as "ข้อบังคับ มจพ. ฉบับที่ 2 พ.ศ. 2568").
- `ประกาศแนวปฏิบัติในการเบิกจ่ายค่าตอบแทนพ...` — the practice announcement governing how the compensation is disbursed.

When changing the eligibility rules, quota numbers, or rate tables in `index.html`, check these PDFs rather than guessing — the JS constants are a direct transcription of tables in these documents.

## Architecture (all inside `index.html`)

The page is a client-only 3-step wizard with no backend calls yet. Comments in the script mark where a real backend belongs (search for `MOCK DATA` and `google.script.run`) — this is built to eventually run as a Google Apps Script web app, so backend integration should follow that pattern (`google.script.run.withSuccessHandler(...)`) rather than `fetch`.

**Steps**, toggled by `goStep(n)` which shows/hides `#step-1/2/3` and rebuilds derived views:
1. **ข้อมูลส่วนตัว (Profile)** — position, staff type, appointment date, field group (`sci_tech` vs `biz_social`), admin role, fiscal year/round. Read via `getCtx()`.
2. **ผลงานทางวิชาการ (Publications)** — a repeating list backed by the `pubs` array; each item has a tier (Scopus quartile / TCI / patent / artwork / etc.), author role(s), and per-item file attachments.
3. **ตรวจสอบและยื่น (Review & submit)** — read-only summary, required-attachment checklist, certification checkboxes gating the submit button.

**Eligibility engine** (`computeEligibility`, `CRITERIA`, `TIER_TO_CRIT`):
- `CRITERIA[fieldGroup][outputType][position]` → `[minQuartileOrAny, regularTarget, adminTarget]`. `fieldGroup` is `sci_tech` or `biz_social`; a `null` entry means that output type doesn't count for that position.
- Each publication contributes `count / target` to its output-type bucket; total score across buckets must reach `>= 1.0` to pass (a "mixed portfolio" — e.g. partial Scopus + partial patents can combine to pass).
- Scopus entries additionally have a minimum quartile gate (`qRank` / `minQRank`) — a pub below the position's required quartile is rejected with a warning instead of counting.
- `TIER_TO_CRIT` maps the publication's selected dropdown tier (e.g. `scopus_q1`, `book_chapter`) to the output-type key used in `CRITERIA`.

**Attachments**: `ATTACH_RULES` is a declarative list of `{key, label, required, show(pub)}` rules; `buildAttachZone`/`refreshAttach` render per-publication attachment slots based on the pub's tier/section/role, matching the checklist in "เอกสารแนบ ๑ หน้า 3" of the regulation.

**Sidebar personal criteria panel** (`buildCriteriaPanel`, `getPersonalCriteria`, `countForCrit`): shows the user's own quota targets and live progress per output type, independent of the pass/fail scorecard, plus the compensation rate table (`RATE_TABLE`/`RATE_NUMERIC`, split 50% automatic / 50% workload-based).

**State**: a single global `pubs` array (each item: `{id, section, title, authors, venue, volume, pages, year, tier, role[], doi, files{}}`) plus a `uid` counter. All rendering is done by re-generating `innerHTML` strings from this state (no framework, no virtual DOM) — mutate `pubs`, then call the relevant `render*`/`update*`/`build*` function to reflect it.

**Document preview/export**: `openPreview()` renders a mock memo (`บันทึกข้อความ`) into a modal from current form state; `dlDocx()` is a stub that just toasts — actual .docx generation is expected to be implemented server-side (Apps Script) later.
