"""
HBS HIMS Lab Bill Automation by Muhammad Imran  —  CONFIRMED WORKING VERSION
═══════════════════════════════════════════════════════════
This version was built and verified step-by-step through live
paste-in-terminal tests on the real HIMS portal. Every behavior
below has been tested against real patients before being written
into this file.

CONFIRMED FLOW (updated — now uses "Summarize Bill" instead of the
old Print → Unpaid flow):
1. Login → Front Desk → Final Settlement
2. Click Patients → toggle the "History" filter ON (so discharged/
   history patients are included) → search by MRN (not name)
   - MRN is unique per patient, so this never returns 2 patients
     with the same identity — no more "wrong card" problem.
   - Before clicking any card, we confirm the MRN is present in
     that card's own text. Only then do we click it.
3. Wait for "Admission History" / "Summarize Bill" to appear, then
   click the "Summarize Bill" link — confirmed to always open a
   new browser tab with the invoice.
4. On the new Summarize Bill invoice tab:
   - Scrape "Admission No" directly from the invoice text
     (e.g. "Admission No: 44574" → 44574). This is the number
     used in the final Excel report, NOT the MRN.
   - Check whether a "LabInvestigation" section exists at all.
     - If not: mark this patient as "no_labs", move to next patient.
     - If yes: extract every LAB-prefixed row between
       "LabInvestigation" and the next section header.
   - This table format has exactly 5 columns per lab row (Service
     ID | Services | Consultant [blank] | Qty | Amount) — no
     separate Rate column, so Rate is derived as Amount / Qty.
     Confirmed on real patients 26046988 and 26048837.
5. Merge duplicate test names (same test appearing more than once
   on one invoice) — sums their Qty and Amount into a single row.
   Confirmed working on a real patient with 2 repeated tests.
6. Close invoice tab → close any stray tabs → return to Final
   Settlement (fresh navigate) → next patient.
   - Confirmed fix: without closing stray tabs AND re-navigating to
     Final Settlement after each patient, the 2nd+ patient's
     "Patients" button click fails with a "subtree intercepts
     pointer events" error (leftover Patient List panel from the
     previous patient blocks it).
7. Build the final Excel report. Column label stays "Admission No"
   even though matching itself is done via MRN — the report shows
   the real Admission No scraped from each invoice.

EXCEL INPUT FORMAT (confirmed):
- Column B → MRN
- Column C → Patient Name
(Column A is not read/used.)

WINDOWS NOTES:
- Sets the correct asyncio event loop policy for Windows so
  Playwright's browser subprocess works correctly.
- All paths use pathlib.Path (works on both Windows and Mac/Linux).
- Run the same way as before:
      python server.py
  or directly:
      python automation.py uploads\\admissions.xlsx "June 26"
"""
import asyncio
import os
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

# ── WINDOWS FIX: correct event loop policy for Playwright subprocesses ──────
if sys.platform.startswith("win"):
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

PORTAL_URL   = "https://hbs.chealth.pk"
ADMIN_ID     = os.getenv("HIMS_USER", "7786")
ADMIN_PASS   = os.getenv("HIMS_PASS",  "randarusy")
DOWNLOAD_DIR = Path("downloads")
OUTPUT_FILE  = Path("output") / "HBS_Lab_Bill.xlsx"
DOWNLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_FILE.parent.mkdir(exist_ok=True)

progress_cb = None


def log(msg):
    print(msg, flush=True)
    if progress_cb:
        progress_cb(msg)


def load_admission_excel(filepath: str) -> list[dict]:
    """
    Reads the input Excel file.
    Column B → MRN
    Column C → Patient Name
    (Column A is ignored, as instructed.)
    """
    df = pd.read_excel(filepath, dtype=str, header=0)
    col_names = list(df.columns)
    mrn_col  = col_names[1]   # Column B
    name_col = col_names[2]   # Column C

    patients = []
    for _, row in df.iterrows():
        mrn  = str(row[mrn_col]).strip()
        name = str(row[name_col]).strip()
        if (mrn and mrn.lower() not in ['nan', ''] and
                name and name.lower() not in ['nan', 'patient name', '']):
            patients.append({"mrn": mrn, "name": name})
    log(f"✅ Loaded {len(patients)} patients from Excel (MRN + Name)")
    return patients


async def login(page):
    log("🔐 Logging in ...")
    await page.goto(PORTAL_URL, wait_until="networkidle", timeout=30000)
    await page.wait_for_timeout(2000)
    await page.fill('input[type="text"]', ADMIN_ID)
    await page.fill('input[type="password"]', ADMIN_PASS)
    await page.click('button[type="submit"]')
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(3000)
    log("✅ Logged in")


async def wait_for_page_ready(page):
    """
    Wait for the ngx-ui-loader spinner to fully disappear before proceeding.
    This prevents the 'ngx-overlay intercepts pointer events' error that
    happens when we try to click the Patients button while the page is
    still loading after returning from a previous patient's invoice tab.
    Confirmed fix via live testing on 4 real patients in sequence.
    """
    try:
        await page.wait_for_selector(
            ".ngx-overlay.loading-foreground, .ngx-overlay.loading-background",
            state="hidden",
            timeout=15000
        )
    except Exception:
        pass
    await page.wait_for_timeout(800)


async def go_to_final_settlement(page):
    log("📋 Opening Front Desk → Final Settlement ...")
    await page.click('text=Front Desk')
    await page.wait_for_timeout(2000)
    links = await page.locator('a:has-text("Final Settlement")').all()
    for link in links:
        if await link.is_visible():
            await link.click()
            break
    await page.wait_for_timeout(2000)
    await wait_for_page_ready(page)
    log("✅ Final Settlement opened")


async def extract_lab_tests(invoice_page) -> tuple[list, float]:
    """
    Extract ALL LabInvestigation rows from the Unpaid invoice page.
    - Scrolls fully (top then bottom) to make sure all rows are loaded
    - Detects lab section start and end accurately
    - Handles rows that don't have exactly 7 columns (flexible parser)
    - Merges duplicate test names: sums qty and amount
      (confirmed working on real repeated tests: Blood CBC ×3,
      BLOOD FOR C/S ×2 → correctly merged into single rows)
    """
    raw_tests = []

    try:
        await invoice_page.evaluate("window.scrollTo(0, 0)")
        await invoice_page.wait_for_timeout(1000)
        await invoice_page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        await invoice_page.wait_for_timeout(2000)
        await invoice_page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        await invoice_page.wait_for_timeout(1500)

        rows = await invoice_page.locator("table tr").all()
        log(f"  📋 Total table rows found: {len(rows)}")
        in_lab = False

        END_SECTIONS = [
            "radiology", "miscservice", "generalmedicine",
            "otmedicine", "pharmacy", "consultation",
            "procedurecharges", "roomcharges", "nursingcharges",
        ]

        for row in rows:
            cells = await row.locator("td, th").all()
            texts = []
            for cell in cells:
                t = await cell.inner_text()
                texts.append(t.strip())

            if not any(t for t in texts):
                continue

            all_text = " ".join(texts).lower().replace(" ", "").replace("\n", "")

            if not in_lab:
                col2_text = texts[2].lower().replace(" ", "") if len(texts) > 2 else ""
                if "labinvestigation" in col2_text or "labinvestigation" in all_text:
                    in_lab = True
                    log("  🔬 Lab section detected — start capturing")
                    continue

            if in_lab and len(texts) >= 3:
                col2_clean = texts[2].lower().replace(" ", "")
                if any(s in col2_clean for s in END_SECTIONS):
                    log(f"  🔚 Lab section ended at: {texts[2]}")
                    in_lab = False
                    continue

            if in_lab:
                if len(texts) == 7:
                    service_id = texts[1].upper().strip()
                    if service_id.startswith("LAB"):
                        try:
                            desc     = texts[2].strip()
                            qty_str  = texts[4].replace(",", "").strip()
                            rate_str = texts[5].replace(",", "").strip()
                            amt_str  = texts[6].replace(",", "").strip()

                            qty    = int(qty_str)    if qty_str    else 1
                            rate   = float(rate_str) if rate_str   else 0.0
                            amount = float(amt_str)  if amt_str    else 0.0

                            if desc and amount > 0:
                                raw_tests.append({
                                    "name":   desc,
                                    "rate":   rate,
                                    "qty":    qty,
                                    "amount": amount,
                                })
                                log(f"  ✅ Test captured: {desc} | Qty:{qty} | Amt:{amount:.0f}")
                        except (ValueError, IndexError) as e:
                            log(f"  ⚠️  Parse error on row {texts}: {e}")
                            continue

                elif len(texts) == 5:
                    # "Summarize Bill" (NEW FLOW) table structure — confirmed
                    # from actual invoice HTML:
                    #   [0] Service ID | [1] Services (desc) | [2] Consultant
                    #   (empty) | [3] Qty | [4] Amount
                    # Ye table "Rate" column deta hi nahi, isliye Rate ko
                    # Amount / Qty se derive karte hain (confirmed on real
                    # patients: 26046988, 26048837).
                    service_id = texts[0].upper().strip()
                    if service_id.startswith("LAB"):
                        try:
                            desc = texts[1].strip()
                            qty_str = texts[3].replace(",", "").strip()
                            amt_str = texts[4].replace(",", "").strip()

                            qty    = int(qty_str)   if qty_str else 1
                            amount = float(amt_str) if amt_str else 0.0
                            rate   = round(amount / qty, 2) if qty else amount

                            if desc and amount > 0:
                                raw_tests.append({
                                    "name":   desc,
                                    "rate":   rate,
                                    "qty":    qty,
                                    "amount": amount,
                                })
                                log(f"  ✅ Test captured (summarize-bill): {desc} | Qty:{qty} | Rate:{rate} | Amt:{amount:.0f}")
                        except (ValueError, IndexError) as e:
                            log(f"  ⚠️  Parse error on row {texts}: {e}")
                            continue

                elif len(texts) > 5:
                    lab_cell_idx = None
                    for idx, t in enumerate(texts):
                        if t.upper().startswith("LAB"):
                            lab_cell_idx = idx
                            break
                    if lab_cell_idx is not None:
                        try:
                            desc_idx = lab_cell_idx + 1
                            desc = texts[desc_idx].strip() if desc_idx < len(texts) else ""
                            amt_str  = texts[-1].replace(",", "").strip()
                            rate_str = texts[-2].replace(",", "").strip() if len(texts) > 2 else "0"
                            qty_str  = texts[-3].replace(",", "").strip() if len(texts) > 3 else "1"
                            qty    = int(qty_str)    if qty_str.isdigit() else 1
                            rate   = float(rate_str) if rate_str   else 0.0
                            amount = float(amt_str)  if amt_str    else 0.0
                            if desc and amount > 0:
                                raw_tests.append({
                                    "name":   desc,
                                    "rate":   rate,
                                    "qty":    qty,
                                    "amount": amount,
                                })
                                log(f"  ✅ Test captured (flex): {desc} | Qty:{qty} | Amt:{amount:.0f}")
                        except (ValueError, IndexError):
                            continue

    except Exception as e:
        log(f"  ⚠️  Extraction error: {e}")

    # ── MERGE DUPLICATES ─────────────────────────────────────────────────
    merged: dict[str, dict] = {}
    for t in raw_tests:
        key = t["name"].lower().strip()
        if key in merged:
            merged[key]["qty"]    += t["qty"]
            merged[key]["amount"] += t["amount"]
            log(f"  🔄 Merged duplicate: {t['name']} → qty now {merged[key]['qty']}")
        else:
            merged[key] = {
                "name":   t["name"],
                "rate":   t["rate"],
                "qty":    t["qty"],
                "amount": t["amount"],
            }

    final_tests = []
    for sr, item in enumerate(merged.values(), start=1):
        final_tests.append({
            "sr":     sr,
            "name":   item["name"],
            "rate":   item["rate"],
            "qty":    item["qty"],
            "amount": item["amount"],
        })

    sub_total = sum(t["amount"] for t in final_tests)

    log(f"  🧪 Raw rows: {len(raw_tests)} | After merge: {len(final_tests)} | Sub Total: {sub_total:.0f}")
    return final_tests, sub_total


def scrape_admission_no(invoice_text: str) -> str:
    """
    Scrapes the real Admission No from the invoice page text.
    Confirmed format: "Admission No:" followed by the number, e.g.
    "Admission No: 44574" → returns "44574"
    """
    match = re.search(r"Admission No[:\s]*([A-Za-z0-9\-]+)", invoice_text)
    return match.group(1).strip() if match else ""


async def close_stray_pages(context, main_page):
    """
    Closes any extra browser tabs except the main Final Settlement page.
    This prevents a stuck/hanging invoice tab from one patient leaking
    into the next patient's processing — each patient now always starts
    with a guaranteed-clean single-tab state.
    """
    for p in list(context.pages):
        if p != main_page and not p.is_closed():
            try:
                await p.close()
                log("  🧹 Closed a leftover tab from a previous step")
            except Exception:
                pass


async def process_patient(page, patient: dict, context) -> dict:
    mrn  = patient["mrn"].strip()
    name = patient["name"].strip()
    log(f"\n👤 Processing: {name}  |  MRN: {mrn}")

    # ── Cleanup BEFORE starting — guarantees a clean slate even if the
    # previous patient timed out or errored with a tab left open.
    await close_stray_pages(context, page)

    result = {
        "patient_name": name,
        "mrn":          mrn,
        "admission_no": "",
        "status":       "pending",
        "tests":        [],
        "sub_total":    0.0,
    }

    try:
        # ── Step 1: Wait for page ready, then click Patients, search by MRN ─
        # Wait for ngx-ui-loader spinner to disappear before clicking —
        # confirmed fix for "ngx-overlay intercepts pointer events" error
        # that blocked Patients button click between sequential patients.
        await wait_for_page_ready(page)

        await page.click('button:has-text("Patients")', timeout=15000)
        await page.wait_for_timeout(1500)

        # Confirmed step from quicktest_summary.py: toggle the "History"
        # filter ON before searching, so history/discharged patients are
        # included in the search results.
        await page.evaluate("""
        () => {
            let els = Array.from(document.querySelectorAll('*')).filter(
                el => el.textContent.trim() === 'History' && el.offsetParent !== null
            );
            for (let el of els) {
                let parent = el.parentElement;
                if(parent){
                    let toggle = parent.querySelector('label,input[type="checkbox"],button');
                    if(!toggle) toggle = parent.previousElementSibling;
                    if(toggle){
                        toggle.click();
                        return;
                    }
                    parent.click();
                }
            }
        }
        """)
        await wait_for_page_ready(page)

        await page.fill('input[placeholder="Search patient here"]', mrn)
        await page.wait_for_timeout(500)
        await page.click('button:has-text("Search")')
        await page.wait_for_timeout(2500)

        # ── Step 2: Confirm MRN is in the card's OWN text BEFORE clicking ─
        # (Confirmed approach: MRN is shown directly inside each result
        # card, so we verify the match before ever clicking anything.)
        candidates = await page.locator(f'text={mrn}').all()
        confirmed_card = None
        for el in candidates:
            if not await el.is_visible():
                continue
            try:
                own_text = await el.inner_text()
            except Exception:
                continue
            if mrn in own_text:
                confirmed_card = el
                break

        if confirmed_card is None:
            log(f"  ❌ No card found with MRN {mrn} for {name}")
            result["status"] = "not_found"
            await go_to_final_settlement(page)
            return result

        log(f"  ✅ Confirmed card matches MRN {mrn} — clicking")
        await confirmed_card.click()
        await page.wait_for_timeout(2500)

        # ── Step 3: Wait for patient data, then click "Summarize Bill" ──────
        # (Confirmed NEW FLOW, replaces the old Print → Unpaid flow.
        # "Summarize Bill" always opens a new tab — confirmed via live
        # testing on real patients.)
        await page.wait_for_selector("text=Admission History", state="visible", timeout=30000)
        await page.wait_for_selector("text=Summarize Bill", state="visible", timeout=30000)
        await wait_for_page_ready(page)

        summary_btn = page.locator(
            "#kt_content > div.d-flex.flex-column-fluid.main-content > div > div > div > div "
            "> app-admission-history > div.card > div > div > table > tbody > tr > td:nth-child(8) > a"
        )
        await summary_btn.wait_for(state="visible", timeout=30000)
        await summary_btn.scroll_into_view_if_needed()
        await page.wait_for_timeout(1000)

        async with context.expect_page(timeout=15000) as new_page_info:
            await summary_btn.click(force=True)
        log("  🖱️  Summarize Bill clicked (new tab opened)")

        invoice_page = await new_page_info.value

        # ── Give the invoice page real time to load ──────────────────────────
        # CONFIRMED via live testing on real patients:
        # - Some pages take 40-60s on first (cold cache) load
        # - Tables appear in the DOM EARLY but with empty/skeleton content
        #   (e.g. table exists with only 62 chars at 8s, then fills to 908
        #   chars with real lab data at 12s — confirmed on real patient)
        # - Stopping at count()>0 caused "no labs" on patients whose
        #   LabInvestigation section rendered slightly after the table
        #   skeleton appeared — this was the real bug causing inconsistency
        # FIX: poll until at least one table has >200 chars of real content,
        # not just until any table exists in the DOM. Budget: 120 seconds.
        table_loaded = False
        wait_start = asyncio.get_event_loop().time()
        max_wait_seconds = 120
        while (asyncio.get_event_loop().time() - wait_start) < max_wait_seconds:
            try:
                all_tables = await invoice_page.locator("table").all()
                if len(all_tables) > 0:
                    for t in all_tables:
                        try:
                            t_text = await t.inner_text()
                            if len(t_text) > 200:
                                table_loaded = True
                                break
                        except Exception:
                            continue
            except Exception:
                pass
            if table_loaded:
                break
            await invoice_page.wait_for_timeout(3000)

        elapsed = round(asyncio.get_event_loop().time() - wait_start, 1)

        if not table_loaded:
            log(f"  ❌ Invoice page never loaded real content after {elapsed}s")
            result["status"] = "load_failed"
            await invoice_page.close()
            await close_stray_pages(context, page)
            await go_to_final_settlement(page)
            return result

        log(f"  ✅ Invoice table loaded after {elapsed}s")

        await invoice_page.wait_for_timeout(1500)

        # inner_text on "body" can itself hang if called too early (seen
        # during testing) — give it an explicit bounded timeout and retry
        # once if needed, now that we already know the table exists.
        try:
            invoice_text = await invoice_page.inner_text("body", timeout=15000)
        except PWTimeout:
            log("  ⏳ Body text read timed out once, retrying...")
            await invoice_page.wait_for_timeout(2000)
            invoice_text = await invoice_page.inner_text("body", timeout=15000)

        # ── Step 5: Scrape Admission No from the invoice ────────────────────
        admission_no = scrape_admission_no(invoice_text)
        result["admission_no"] = admission_no
        log(f"  🔢 Scraped Admission No: {admission_no or '(not found)'}")

        # ── Step 6: Check if a LabInvestigation section exists at all ──────
        has_lab_section = "labinvestigation" in invoice_text.lower().replace(" ", "")

        if not has_lab_section:
            log(f"  ℹ️  No LabInvestigation section found — no labs for this patient")
            result["status"] = "no_labs"

            safe_name = re.sub(r'[\\/:*?"<>|]', "_", name)
            pdf_path = str(DOWNLOAD_DIR / f"{mrn}_{safe_name}.pdf")
            try:
                await invoice_page.pdf(
                    path=pdf_path, format="A4", print_background=True,
                    margin={"top": "10mm", "bottom": "10mm", "left": "10mm", "right": "10mm"}
                )
            except Exception as e:
                log(f"  ⚠️  PDF save failed: {e}")

            await invoice_page.close()
            await go_to_final_settlement(page)
            return result

        # ── Step 7: Extract lab tests (with duplicate merging) ─────────────
        tests, sub_total = await extract_lab_tests(invoice_page)

        # ── Step 8: Save PDF ─────────────────────────────────────────────────
        safe_name = re.sub(r'[\\/:*?"<>|]', "_", name)   # Windows-safe filename
        pdf_path = str(DOWNLOAD_DIR / f"{mrn}_{safe_name}.pdf")
        try:
            await invoice_page.pdf(
                path=pdf_path, format="A4", print_background=True,
                margin={"top": "10mm", "bottom": "10mm", "left": "10mm", "right": "10mm"}
            )
            log(f"  📄 PDF saved")
        except Exception as e:
            log(f"  ⚠️  PDF save failed: {e}")

        await invoice_page.close()

        result["tests"]     = tests
        result["sub_total"] = sub_total
        result["status"]    = "success"

    except PWTimeout as e:
        log(f"  ⏱️  Timeout: {e}")
        result["status"] = "timeout"
    except Exception as e:
        log(f"  ❌ Error: {e}")
        result["status"] = "error"

    # ── Cleanup AFTER, regardless of success/failure — closes any
    # invoice tab that might still be open (e.g. if the table never
    # appeared and we continued anyway, or an error happened mid-step).
    await close_stray_pages(context, page)

    await go_to_final_settlement(page)
    return result


def build_lab_bill_excel(all_results: list[dict], month_label: str = ""):
    wb = Workbook()
    ws = wb.active
    ws.title = "Lab Bill"

    thin   = Side(style="thin",   color="000000")
    thick  = Side(style="medium", color="000000")
    thin_b = Border(left=thin, right=thin, top=thin, bottom=thin)
    H_FILL    = PatternFill("solid", start_color="1F4E79")
    SUBH_FILL = PatternFill("solid", start_color="BDD7EE")
    ALT_FILL  = PatternFill("solid", start_color="F2F9FF")
    WHITE     = PatternFill("solid", start_color="FFFFFF")
    TOTAL_F   = PatternFill("solid", start_color="DEEAF1")
    NOLABS_F  = PatternFill("solid", start_color="FFF3E0")

    def cs(c, bold=False, size=10, color="000000", fill=None, align="left", border=None):
        c.font = Font(name="Calibri", bold=bold, size=size, color=color)
        c.alignment = Alignment(horizontal=align, vertical="center")
        if fill:   c.fill = fill
        if border: c.border = border

    for col, w in zip(["A", "B", "C", "D", "E"], [8, 30, 10, 6, 12]):
        ws.column_dimensions[col].width = w

    row = 1
    RED_F = PatternFill("solid", start_color="FFE0E0")

    for res in all_results:
        name         = res.get("patient_name", "")
        admission_no = res.get("admission_no", "") or "—"
        tests        = res.get("tests", [])
        sub_total    = res.get("sub_total", 0.0)
        status       = res.get("status", "")
        bf           = (NOLABS_F if status == "no_labs"
                         else RED_F if status == "load_failed"
                         else WHITE)

        ws.merge_cells(f"A{row}:E{row}")
        c = ws[f"A{row}"]; c.value = "HBS GENERAL HOSPITAL"
        cs(c, bold=True, size=13, color="FFFFFF", fill=H_FILL, align="center")
        ws.row_dimensions[row].height = 22; row += 1

        ws.merge_cells(f"A{row}:E{row}")
        c = ws[f"A{row}"]; c.value = f"Laboratory Investigation Month of {month_label}"
        cs(c, size=10, fill=SUBH_FILL, align="center")
        ws.row_dimensions[row].height = 16; row += 1

        ws.merge_cells(f"A{row}:E{row}")
        c = ws[f"A{row}"]; c.value = "Invoice"
        cs(c, bold=True, size=11, align="center")
        ws.row_dimensions[row].height = 16; row += 1

        name_suffix = (
            "   (No Labs)" if status == "no_labs"
            else "   ⚠ COULD NOT VERIFY — PLEASE CHECK MANUALLY" if status == "load_failed"
            else ""
        )
        ws.merge_cells(f"A{row}:B{row}")
        c = ws[f"A{row}"]
        c.value = f"Patient Name:  {name}{name_suffix}"
        cs(c, bold=True, size=10, fill=bf,
           color="C05621" if status == "no_labs" else "C00000" if status == "load_failed" else "000000")
        ws.merge_cells(f"C{row}:E{row}")
        c2 = ws[f"C{row}"]; c2.value = f"Admission No.  {admission_no}"
        cs(c2, bold=True, size=10, fill=bf)
        ws.row_dimensions[row].height = 18; row += 1

        ws.merge_cells(f"A{row}:E{row}")
        c = ws[f"A{row}"]; c.value = "Laboratory Investigation"
        cs(c, bold=True, size=10, color="FFFFFF", fill=H_FILL, align="center")
        ws.row_dimensions[row].height = 16; row += 1

        for col, h, algn in zip(["A", "B", "C", "D", "E"],
                                 ["Sr.No.", "Investigation Name", "Rate", "Q", "Amount"],
                                 ["center", "left", "center", "center", "right"]):
            c = ws[f"{col}{row}"]; c.value = h
            cs(c, bold=True, size=10, color="FFFFFF", fill=H_FILL, align=algn, border=thin_b)
        ws.row_dimensions[row].height = 16; row += 1

        if tests:
            for i, test in enumerate(tests):
                rf = ALT_FILL if i % 2 == 0 else WHITE
                for col, val, algn in zip(
                    ["A", "B", "C", "D", "E"],
                    [test["sr"], test["name"], test["rate"], test["qty"], test["amount"]],
                    ["center", "left", "right", "center", "right"]
                ):
                    c = ws[f"{col}{row}"]; c.value = val
                    cs(c, size=10, fill=rf, align=algn, border=thin_b)
                    if col in ["C", "E"]: c.number_format = "#,##0"
                ws.row_dimensions[row].height = 15; row += 1
        else:
            ws.merge_cells(f"A{row}:E{row}")
            c = ws[f"A{row}"]
            if status == "no_labs":
                c.value = "— No Labs —"
                cs(c, size=10, color="888888", align="center", fill=WHITE)
            elif status == "load_failed":
                c.value = "— ⚠ Invoice page did not load — run this patient again to verify —"
                cs(c, bold=True, size=10, color="C00000", align="center", fill=RED_F)
            else:
                c.value = "— Could Not Be Processed —"
                cs(c, size=10, color="888888", align="center", fill=WHITE)
            ws.row_dimensions[row].height = 15; row += 1

        ws.merge_cells(f"A{row}:D{row}")
        c = ws[f"A{row}"]; c.value = "Sub Total Of Laboratory"
        cs(c, bold=True, size=10, fill=TOTAL_F, align="right",
           border=Border(top=thick, bottom=thick))
        c2 = ws[f"E{row}"]; c2.value = sub_total
        cs(c2, bold=True, size=10, fill=TOTAL_F, align="right",
           border=Border(top=thick, bottom=thick, left=thin, right=thin))
        c2.number_format = "#,##0"
        ws.row_dimensions[row].height = 16; row += 1

        ws.merge_cells(f"A{row}:E{row}")
        c = ws[f"A{row}"]; c.value = "CHECKED BY"
        cs(c, size=9, color="666666", align="center")
        ws.row_dimensions[row].height = 14; row += 1

        ws.merge_cells(f"B{row}:D{row}")
        c = ws[f"B{row}"]; c.value = "_______________________________"
        cs(c, size=9, align="center")
        ws.row_dimensions[row].height = 14; row += 1

        ws.row_dimensions[row].height = 10; row += 1

    # ── Summary sheet ────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    for col, w in zip(["A", "B", "C", "D", "E", "F", "G"], [6, 12, 12, 24, 12, 14, 14]):
        ws2.column_dimensions[col].width = w

    ws2.merge_cells("A1:G1")
    c = ws2["A1"]; c.value = f"HBS GENERAL HOSPITAL — Lab Bill Summary — {month_label}"
    c.font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    c.fill = H_FILL; c.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 26

    for col, h in zip(["A", "B", "C", "D", "E", "F", "G"],
                       ["Sr.", "MRN", "Admission No.", "Patient Name", "Lab Tests", "Sub Total (PKR)", "Status"]):
        c = ws2[f"{col}2"]; c.value = h
        c.font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
        c.fill = H_FILL; c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_b
    ws2.row_dimensions[2].height = 18

    STATUS_LABELS = {
        "success":     ("✅ Done",            "00A86B"),
        "no_labs":     ("ℹ️  No Labs",         "C05621"),
        "load_failed": ("⚠️  Recheck Manually", "C00000"),
        "not_found":   ("❌ Not Found",        "C00000"),
        "timeout":     ("⏱ Timeout",          "888888"),
        "error":       ("❌ Error",            "C00000"),
    }

    for i, res in enumerate(all_results):
        r = i + 3
        fill = ALT_FILL if i % 2 == 0 else WHITE
        lbl, clr = STATUS_LABELS.get(res.get("status", ""), ("—", "000000"))
        for col, val, algn in zip(
            ["A", "B", "C", "D", "E", "F", "G"],
            [i + 1, res.get("mrn", ""), res.get("admission_no", "") or "—",
             res.get("patient_name", ""), len(res.get("tests", [])),
             res.get("sub_total", 0.0), lbl],
            ["center", "left", "left", "left", "center", "right", "center"]
        ):
            c = ws2[f"{col}{r}"]; c.value = val; c.fill = fill; c.border = thin_b
            c.font = Font(name="Calibri", size=10,
                          color=clr if col == "G" else "000000", bold=(col == "G"))
            c.alignment = Alignment(horizontal=algn, vertical="center")
            if col == "F": c.number_format = "#,##0"
        ws2.row_dimensions[r].height = 15

    tr = len(all_results) + 3
    ws2.merge_cells(f"A{tr}:E{tr}")
    c = ws2[f"A{tr}"]; c.value = "GRAND TOTAL"
    c.font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
    c.fill = PatternFill("solid", start_color="C00000")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = Border(top=thick, bottom=thick)
    c2 = ws2[f"F{tr}"]; c2.value = f"=SUM(F3:F{tr-1})"
    c2.font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
    c2.fill = PatternFill("solid", start_color="C00000")
    c2.number_format = "#,##0"
    c2.alignment = Alignment(horizontal="right", vertical="center")
    c2.border = Border(top=thick, bottom=thick)
    ws2[f"G{tr}"].fill = PatternFill("solid", start_color="C00000")
    ws2[f"G{tr}"].border = Border(top=thick, bottom=thick)
    ws2.row_dimensions[tr].height = 22

    wb.save(str(OUTPUT_FILE))
    log(f"\n✅ Excel saved → {OUTPUT_FILE}")


async def run_pipeline(admission_excel: str, month_label: str = "", cb=None):
    global progress_cb
    progress_cb = cb
    patients = load_admission_excel(admission_excel)
    results  = []

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        context = await browser.new_context(accept_downloads=True)
        page    = await context.new_page()

        await login(page)
        await go_to_final_settlement(page)

        for i, patient in enumerate(patients):
            log(f"\n[{i+1}/{len(patients)}]")
            res = await process_patient(page, patient, context)
            results.append(res)
            await asyncio.sleep(1.0)

        await browser.close()

    build_lab_bill_excel(results, month_label)

    success      = sum(1 for r in results if r["status"] == "success")
    no_labs      = sum(1 for r in results if r["status"] == "no_labs")
    load_failed  = sum(1 for r in results if r["status"] == "load_failed")
    failed       = len(results) - success - no_labs - load_failed
    total    = sum(r.get("sub_total", 0) for r in results)
    log(f"\n{'='*50}")
    log(f"✅ Done | {success} with labs | {no_labs} no labs | {load_failed} need recheck | {failed} failed")
    log(f"💰 Grand Total Lab: PKR {total:,.0f}")
    return results


if __name__ == "__main__":
    excel = sys.argv[1] if len(sys.argv) > 1 else str(Path("uploads") / "admissions.xlsx")
    month = sys.argv[2] if len(sys.argv) > 2 else "June 26"
    asyncio.run(run_pipeline(excel, month))