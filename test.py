# pulse_playwright.py
import asyncio
import os
import re
import string
import traceback
from typing import List, Optional, Dict

import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError, Page

# -------------------------
# Configuration
# -------------------------
MAX_CONCURRENT = 4  # number of parallel pages/workers (tune for speed)
BROWSER_HEADLESS = False  # set True to run headless
AUTOSAVE_PATH = r"C:\PULSE_Auto\PULSE_AUTO.xlsx"  # temp autosave used similarly to your old script

# -------------------------
# Helper types
# -------------------------
PartResult = Dict[str, object]  # structure to return row-data for a single part (flexible)


# -------------------------
# Main App (keeps Tkinter UI + openpyxl behavior)
# -------------------------
class App:
    def __init__(self):
        # Tkinter setup (kept like original)
        self.root = tk.Tk()
        self.root.title("PULSE")
        self.root.withdraw()

        # State variables
        self.Part_nos: List[str] = []
        self.excel_file: Optional[str] = None
        self.folder_path: Optional[str] = None
        self.filename: Optional[str] = None
        self.workbook = None
        self.sheet1 = None
        self.sheet2 = None
        self.sheet3 = None
        self.row_index = 1

        # Excel styles (mirrors your original style attributes)
        self.green = PatternFill(start_color='c6efce', end_color='c6efce', fill_type='solid')
        self.red = PatternFill(start_color='ffc7ce', end_color='ffc7ce', fill_type='solid')
        self.blue = PatternFill(start_color='305496', end_color='305496', fill_type='solid')
        self.Thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                  top=Side(style='thin'), bottom=Side(style='thin'))
        self.Thick_border = Border(left=Side(style='thick'), right=Side(style='thick'),
                                   top=Side(style='thick'), bottom=Side(style='thick'))
        self.bold_font = Font(bold=True)
        self.Center_align = Alignment(horizontal='center', vertical='center')

        # Async locks
        self.workbook_lock = asyncio.Lock()  # ensure only one coroutine writes to workbook at a time

    # -------------------------
    # Excel and input initialization (similar to Basic)
    # -------------------------
    def Basic(self):
        """Open workbook, clear/create PULSE DATA, and read Part_nos from INPUTS column B (starting row 2)."""
        # Let user pick the file via dialog (keeps UX similar to your tkinter usage)
        file_path = filedialog.askopenfilename(title="Select input Excel file",
                                               filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xltx;*.xltm")])
        if not file_path:
            messagebox.showerror("No file", "No Excel file selected.")
            self.root.quit()
            return

        self.excel_file = file_path
        self.folder_path, self.filename = os.path.split(file_path)

        # Load workbook
        self.workbook = load_workbook(file_path)
        # If PULSE Data exists in old variants of your code: remove and recreate (keeps behavior)
        if "PULSE DATA" in [s.upper() for s in self.workbook.sheetnames]:
            # remove any case variant
            for name in list(self.workbook.sheetnames):
                if name.upper() == "PULSE DATA":
                    std = self.workbook[name]
                    self.workbook.remove(std)
        # Create new "PULSE DATA"
        self.workbook.create_sheet("PULSE DATA")
        self.sheet1 = self.workbook["INPUTS"]
        self.sheet2 = self.workbook["PULSE DATA"]
        # If original had "Report Data" sheet, keep reference (if present)
        if "Report Data" in self.workbook.sheetnames:
            self.sheet3 = self.workbook["Report Data"]
        else:
            self.sheet3 = None

        # Adjust columns widths (mirror from original)
        widths = [20, 20, 20, 20, 10, 20, 10, 10, 20, 20, 26, 10, 20, 20]
        for col_letter, w in zip(string.ascii_uppercase[:len(widths)], widths):
            self.sheet2.column_dimensions[col_letter].width = w

        # Read part numbers from column B starting at row 2 (Selenium code used sheet1['B'][1:])
        self.Part_nos = []
        for cell in self.sheet1['B'][1:]:
            if cell.value is None:
                break
            self.Part_nos.append(str(cell.value).strip())

        # Start row index at 1 (like your code). We'll append rows using this shared counter.
        self.row_index = 1

    # -------------------------
    # Core: run Playwright and manage parallel workers
    # -------------------------
    async def run_playwright_parallel(self):
        """
        Creates a single browser and context, then runs multiple page workers concurrently.
        For each part number, we open a new page (tab) and run the same logic as your original script.
        Workbook writes are synchronized with self.workbook_lock to avoid corruption.
        """
        if not self.Part_nos:
            print("No parts to process.")
            return

        playwright = await async_playwright().start()
        # Launch Firefox (maps to webdriver.Firefox() in Selenium)
        browser = await playwright.firefox.launch(headless=BROWSER_HEADLESS)
        context = await browser.new_context()

        # Open initial page (like self.driver.get('https://hihi.com'))
        # We'll create a "starter" page to handle any initial click that opens new window in original.
        starter_page = await context.new_page()
        await starter_page.goto("https://hihi.com", timeout=60000)

        # Try to click 'Continue' input if present (Selenium used find_elements and click)
        try:
            # Playwright locator: input[value="Continue"]
            continue_locator = starter_page.locator('input[value="Continue"]')
            if await continue_locator.count() > 0:
                try:
                    await continue_locator.first.click(timeout=5000)
                except PlaywrightTimeoutError:
                    pass
        except Exception:
            pass

        # After possible click, wait for network idle (represents page load)
        try:
            await starter_page.wait_for_load_state('networkidle', timeout=30000)
        except Exception:
            # ignore timeouts; continue
            pass

        # If the original flow opens a new window, Playwright's context.pages will include it.
        # We'll use the last page for worker templates. But we will create fresh pages per worker.
        # Use semaphore to limit concurrency
        semaphore = asyncio.Semaphore(MAX_CONCURRENT)

        # Create a list of tasks
        tasks = []
        for part_no in self.Part_nos:
            task = asyncio.create_task(self._worker_process_part(context, part_no, semaphore))
            tasks.append(task)

        # Wait for all tasks to finish
        results = await asyncio.gather(*tasks, return_exceptions=True)

        # Save workbook final and cleanup
        async with self.workbook_lock:
            try:
                self.workbook.save(os.path.join(self.folder_path, self.filename))
            except Exception:
                # fallback: try autosave path
                try:
                    self.workbook.save(AUTOSAVE_PATH)
                except Exception:
                    traceback.print_exc()

        # close Playwright
        await context.close()
        await browser.close()
        await playwright.stop()

        # open workbook (like os.startfile in your original)
        try:
            os.startfile(os.path.join(self.folder_path, self.filename))
        except Exception:
            # ignore platform issues
            pass

        # print any exceptions from tasks
        for r in results:
            if isinstance(r, Exception):
                print("Task failed:", r)

    # -------------------------
    # Worker: do the automation for a single part number
    # Mirrors your main_code loop body but for one part
    # -------------------------
    async def _worker_process_part(self, context, Part_no: str, semaphore: asyncio.Semaphore) -> Optional[PartResult]:
        """
        Each worker opens its own page and performs the interactions for one Part_no.
        All writes to the workbook are done under self.workbook_lock.
        """
        await semaphore.acquire()
        page: Page = await context.new_page()
        try:
            # -------------------------
            # Navigation & initial clicks
            # -------------------------
            # Equivalent of driver.get(...) and switching windows
            await page.goto("https://hihi.com", timeout=60000)
            # replace sleeps with waits
            try:
                # Try to click #hello and #hilli if present (maps to WebDriverWait(...).click() in your code)
                for sel in ("#hello", "#hilli"):
                    try:
                        locator = page.locator(sel)
                        if await locator.count() > 0:
                            try:
                                await locator.first.click(timeout=5000)
                            except PlaywrightTimeoutError:
                                pass
                    except Exception:
                        pass
            except Exception:
                pass

            # Wait for page to be stable
            try:
                await page.wait_for_load_state("networkidle", timeout=20000)
            except Exception:
                pass

            # -------------------------
            # Prepare initial rows for this part (mirrors how your loop writes "JK", "Def Mef", "JIR" then modifies)
            # We'll collect the changes to write with workbook_lock to avoid collisions
            # -------------------------
            # Prepare a batch of cell updates to write later
            write_updates = []  # list of tuples (cell, value, optional_style)
            # In original code, before searching they saved workbook and wrote A row with JK etc
            write_updates.append((f"A{self.row_index}", "JK"))
            write_updates.append((f"B{self.row_index}", "Def Mef"))
            write_updates.append((f"C{self.row_index}", "JIR"))

            # Apply header styles on columns A-C in this row
            header_cells = [f"{let}{self.row_index}" for let in string.ascii_uppercase[:3]]
            # We'll apply styles when writing to workbook under lock

            # Simulate search interactions (Selenium: find_element, clear, execute_script)
            try:
                search_locator = page.locator("#searchkalu")
                # clear (Playwright fill with empty string)
                if await search_locator.count() > 0:
                    await search_locator.fill("")  # equivalent to clear()
                # execute_script showFile('1) or "showFile('1');" in your code had a small typo - we call the intended value
                # original: self.driver.execute_script("showFile('1);") <- broken quotes; assuming showFile('1')
                try:
                    await page.evaluate("showFile('1');")
                except Exception:
                    # fallback: try with double quotes
                    try:
                        await page.evaluate('showFile("1");')
                    except Exception:
                        pass
            except Exception:
                pass

            # Save/writing initial JK row to workbook
            async with self.workbook_lock:
                # Write the header row and apply styles
                for (cell, value) in write_updates:
                    self.sheet2[cell] = value
                for let in string.ascii_uppercase[:3]:
                    c = self.sheet2[f"{let}{self.row_index}"]
                    c.fill = self.blue
                    c.font = Font(color="FFFFFF", bold=True)
                    c.border = self.Thin_border

                # Save an autosave (like original)
                try:
                    self.workbook.save(AUTOSAVE_PATH)
                except Exception:
                    pass

            # Move row index forward and write Part_no row (same as original)
            # NOTE: Since multiple workers share self.row_index, we must update this under the lock
            async with self.workbook_lock:
                this_row = self.row_index + 1
                self.sheet2[f"A{this_row}"] = Part_no
                # increment shared row_index to reflect reserved rows used by this worker
                # We'll increment by 1 here; later writes will also increment under the lock where needed
                self.row_index = this_row

            # -------------------------
            # Extract def_mef and jir from page (Selenium used .get_attribute('value') on element by ID)
            # Playwright equivalent: locator.input_value() or get_attribute('value')
            # IDs: 'moderbean.defmed' and 'moderbean.jir'
            # Note: Playwright selectors with dots require escaping the dot with \\.
            # -------------------------
            try:
                def_mef_val = await page.locator("#moderbean\\.defmed").input_value(timeout=3000)
            except Exception:
                def_mef_val = ""
            try:
                jir_val = await page.locator("#moderbean\\.jir").input_value(timeout=3000)
            except Exception:
                jir_val = ""

            # write the def_mef and jir into the same row (this_row)
            async with self.workbook_lock:
                self.sheet2[f"B{this_row}"] = def_mef_val
                self.sheet2[f"C{this_row}"] = jir_val
                # style the A-C cells for this row
                for let in string.ascii_uppercase[:3]:
                    c = self.sheet2[f"{let}{this_row}"]
                    c.fill = self.blue
                    c.border = self.Thin_border

            # prepare next row for FC section like original (increment row_index)
            async with self.workbook_lock:
                self.row_index += 1
                fc_row = self.row_index  # this will be used to write FC header/title rows

                # Write "FC" cell and style (mirrors original)
                self.sheet2[f"A{fc_row}"] = "FC"
                self.sheet2[f"A{fc_row}"].fill = self.blue
                self.sheet2[f"A{fc_row}"].border = self.Thin_border
                self.sheet2[f"A{fc_row}"].font = Font(color="FFFFFF", bold=True)

                # Merge B:C for FC Title (the original used B{row}:C{row})
                merge_title = f"B{fc_row}:C{fc_row}"
                self.sheet2[f"B{fc_row}"] = "FC Title"
                self.sheet2[f"B{fc_row}"].fill = self.blue
                self.sheet2[f"B{fc_row}"].border = self.Thin_border
                self.sheet2[f"B{fc_row}"].font = Font(color="FFFFFF", bold=True)
                try:
                    self.sheet2.merge_cells(merge_title)
                except Exception:
                    pass

                # Merge D:L for FC Text and set value
                merge_text = f"D{fc_row}:L{fc_row}"
                self.sheet2[f"D{fc_row}"] = "FC Text"
                self.sheet2[f"D{fc_row}"].fill = self.blue
                self.sheet2[f"D{fc_row}"].border = self.Thin_border
                self.sheet2[f"D{fc_row}"].font = Font(color="FFFFFF", bold=True)
                try:
                    self.sheet2.merge_cells(merge_text)
                except Exception:
                    pass

                # Merge M:N for Revised
                merge_rev = f"M{fc_row}:N{fc_row}"
                self.sheet2[f"M{fc_row}"] = "Revised"
                self.sheet2[f"M{fc_row}"].fill = self.blue
                self.sheet2[f"M{fc_row}"].border = self.Thin_border
                self.sheet2[f"M{fc_row}"].font = Font(color="FFFFFF", bold=True)
                try:
                    self.sheet2.merge_cells(merge_rev)
                except Exception:
                    pass

            # increment row_index and then call showFile('4') (mapping to your driver.execute_script("showFile('4');"))
            try:
                await page.evaluate("showFile('4');")
            except Exception:
                try:
                    await page.evaluate('showFile("4");')
                except Exception:
                    pass

            # small wait for UI to render after showFile
            try:
                await page.wait_for_timeout(1000)
            except Exception:
                pass

            # -------------------------
            # FC extraction loop: find inputs that start with name "fcthoda[" (original xpath)
            # Playwright: use locator('xpath=//input[starts-with(@name,"fcthoda[")]')
            # Then for each matched_index try to read fcbean[{i}].fc .value and check if in FC_text_code
            # -------------------------
            FC_text_code = ['1', '2', '3', '4', 'MF', 'AB']
            try:
                fc_inputs = page.locator('xpath=//input[starts-with(@name,"fcthoda[")]')
                count_fc = await fc_inputs.count()
            except Exception:
                count_fc = 0

            Matched_index = 0
            # we'll iterate up to count_fc or a safe max to avoid runaway loops
            for _ in range(count_fc):
                try:
                    # original: exact_fc = driver.find_element(By.NAME, f'fcbean[{Matched_index}].fc')
                    name_selector = f'[name="fcbean[{Matched_index}].fc"]'
                    if await page.locator(name_selector).count() == 0:
                        Matched_index += 1
                        continue
                    FC_code = await page.locator(name_selector).input_value(timeout=2000)
                except Exception:
                    Matched_index += 1
                    continue

                if FC_code in FC_text_code:
                    # Write header for FC code row
                    async with self.workbook_lock:
                        cur = self.row_index
                        self.sheet2[f"A{cur}"] = "FC Code"
                        self.sheet2[f"A{cur}"].fill = self.blue
                        self.sheet2[f"A{cur}"].border = self.Thin_border
                        self.sheet2[f"A{cur}"].font = Font(color="FFFFFF", bold=True)

                        # merge B:C for title and put title
                        merge_TITLE = f"B{cur}:C{cur}"
                        try:
                            fc_title = await page.locator(f'[name="fcBean[{Matched_index}].fcTitle"]').input_value(timeout=2000)
                        except Exception:
                            fc_title = ""
                        self.sheet2[f"B{cur}"] = fc_title
                        try:
                            self.sheet2.merge_cells(merge_TITLE)
                        except Exception:
                            pass

                        # merge D:L for FC text and fill
                        merge_cell_FC = f"D{cur}:L{cur}"
                        # try different possible element names similar to your try/except chain
                        frm = ""
                        try:
                            frm = await page.locator(f'[name="FCText[{Matched_index}].freeFormTxt"]').input_value(timeout=1000)
                        except Exception:
                            try:
                                frm = await page.locator(f'[name="fcText[{Matched_index}].overLengthPArt"]').input_value(timeout=1000)
                            except Exception:
                                try:
                                    frm = await page.locator(f'[name="fcText[{Matched_index}].specNum"]').input_value(timeout=1000)
                                except Exception:
                                    frm = ""
                        self.sheet2[f"D{cur}"] = frm
                        try:
                            self.sheet2.merge_cells(merge_cell_FC)
                        except Exception:
                            pass

                        # revised dates (two cells M and N)
                        Rev_by = ""
                        Rev_date = ""
                        try:
                            Rev_by = await page.locator(f'[name="fcText[{Matched_index}].revisedDatestr"]').input_value(timeout=1000)
                        except Exception:
                            Rev_by = ""
                        try:
                            Rev_date = await page.locator(f'[name="fcText[{Matched_index}].revisedDatestr"]').input_value(timeout=1000)
                        except Exception:
                            Rev_date = ""

                        self.sheet2[f"M{cur}"] = Rev_by
                        self.sheet2[f"N{cur}"] = Rev_date

                        # increment shared row
                        self.row_index += 1
                Matched_index += 1

            # -------------------------
            # WY image check block
            # Original: WY_exist = ["https://hehe.png","https://hihi.png"]
            # Check element with id 'iw_img' and if src in WY_exist then click name 'bt_INDIRECTWY'
            # -------------------------
            WY_exist = ["https://hehe.png", "https://hihi.png"]
            try:
                img_locator = page.locator("#iw_img")
                if await img_locator.count() > 0:
                    WY_exist_img = await img_locator.get_attribute("src")
                else:
                    WY_exist_img = None
            except Exception:
                WY_exist_img = None

            if WY_exist_img in WY_exist:
                # click bt_INDIRECTWY (original code uses find_element(By.NAME,'bt_INDIRECTWY').click())
                try:
                    click_locator = page.locator('[name="bt_INDIRECTWY"]')
                    if await click_locator.count() > 0:
                        await click_locator.first.click(timeout=4000)
                except Exception:
                    pass

                # small pause for UI rendering
                try:
                    await page.wait_for_timeout(1000)
                except Exception:
                    pass

                # Write header row for Indirect/Direct table in sheet
                async with self.workbook_lock:
                    cur = self.row_index
                    header_titles = ["Dir/Indir", "Indir Active", "JK", "Type", "MN/CN",
                                     "MN/CN JK", "WY", "AW", "Okay used on", "Rev", "Derived", "B or C", "Revised"]
                    for col_idx, title in enumerate(header_titles):
                        col = string.ascii_uppercase[col_idx]  # 0->A, 1->B, ...
                        self.sheet2[f"{col}{cur}"] = title
                        # style
                        self.sheet2[f"{col}{cur}"].fill = self.blue
                        self.sheet2[f"{col}{cur}"].font = Font(color="FFFFFF", bold=True)
                        self.sheet2[f"{col}{cur}"].border = self.Thin_border

                    # merge M:N similar to original
                    try:
                        self.sheet2.merge_cells(f"M{cur}:N{cur}")
                    except Exception:
                        pass

                    # increment row
                    self.row_index += 1

                # iterate indirect WY entries
                WYS = ['1', '2', '5']
                Matched_WY = 0
                Matched_index = 0  # used to navigate list items by name indexing
                # For robustness, attempt to loop until no more items or a safe cap
                safe_cap = 200
                entries_found = 0
                while entries_found < safe_cap:
                    # attempt to read value of name f'IndirectWYlist[{Matched_index}].wycode'
                    try:
                        wy_sel = f'[name="IndirectWYlist[{Matched_index}].wycode"]'
                        if await page.locator(wy_sel).count() == 0:
                            # attempt to click next page button similar to your code
                            try:
                                next_page_btn = page.locator('[name="bt_Page"]')
                                if await next_page_btn.count() > 0:
                                    # check if enabled: Playwright does not have is_enabled exactly but we can try click
                                    try:
                                        await next_page_btn.first.click(timeout=2000)
                                        Matched_WY = 0
                                        Matched_index = 0
                                        entries_found += 1
                                        continue
                                    except Exception:
                                        break
                                else:
                                    break
                            except Exception:
                                break
                        WY_id = await page.locator(wy_sel).input_value(timeout=2000)
                    except Exception:
                        break

                    if str(WY_id) in WYS:
                        # Direct/Indirect
                        try:
                            direct_val = await page.locator(f'[name="indirectList[{Matched_WY}].direct"]').input_value(timeout=1000)
                        except Exception:
                            direct_val = ""

                        async with self.workbook_lock:
                            cur = self.row_index
                            self.sheet2[f"A{cur}"] = "Direct" if direct_val == "true" else "Indirect"

                        # status
                        try:
                            Indirect_status = await page.locator(f'[name="indirectList[{Matched_WY}].status"]').input_value(timeout=1000)
                        except Exception:
                            Indirect_status = ""
                        async with self.workbook_lock:
                            if Indirect_status in ["A", "Active"]:
                                self.sheet2[f"B{cur}"] = "Active"
                                self.sheet2[f"B{cur}"].font = self.bold_font

                        # New_JK
                        try:
                            New_JK = await page.locator(f'[name="indirectList[{Matched_index}].partNumber"]').input_value(timeout=1000)
                        except Exception:
                            New_JK = ""
                        async with self.workbook_lock:
                            self.sheet2[f"C{cur}"] = New_JK
                            self.sheet2[f"C{cur}"].fill = self.green

                        # Type_name
                        try:
                            Type_name = await page.locator(f'[name="indirectList[{Matched_index}].typeName"]').input_value(timeout=1000)
                        except Exception:
                            Type_name = ""
                        async with self.workbook_lock:
                            self.sheet2[f"D{cur}"] = Type_name

                        # MN/CN
                        try:
                            MN_CN = await page.locator(f'[name="indirectList[{Matched_index}].mncn"]').input_value(timeout=1000)
                        except Exception:
                            MN_CN = ""
                        async with self.workbook_lock:
                            if MN_CN:
                                self.sheet2[f"E{cur}"] = MN_CN

                        # Old_JK
                        try:
                            Old_JK = await page.locator(f'[name="indirectList[{Matched_index}].oldiePArt"]').input_value(timeout=1000)
                        except Exception:
                            Old_JK = ""
                        async with self.workbook_lock:
                            self.sheet2[f"F{cur}"] = Old_JK
                            # original code assigned fill=self.red incorrectly; we apply red fill
                            self.sheet2[f"F{cur}"].fill = self.red

                        # WY_type, Rework, applicable, Rev_EQ, Auth_Der, B_C, Rev_1, Rev2
                        def safe_get(name_idx, field_name):
                            sel = f'[name="indirectList[{name_idx}].{field_name}"]'
                            try:
                                return page.locator(sel).input_value(timeout=1000)
                            except Exception:
                                return ""

                        WY_type = await safe_get(Matched_index, "wycode")
                        Rework = await safe_get(Matched_index, "rework")
                        applicable = await safe_get(Matched_index, "app")
                        Rev_EQ = await safe_get(Matched_index, "eq")
                        Auth_Der = await safe_get(Matched_index, "authe")
                        B_C = await safe_get(Matched_index, "bc")
                        Rev_1 = await safe_get(Matched_index, "rev1")
                        Rev2 = await safe_get(Matched_index, "rev2")

                        async with self.workbook_lock:
                            self.sheet2[f"G{cur}"] = WY_type
                            self.sheet2[f"H{cur}"] = Rework
                            self.sheet2[f"I{cur}"] = applicable
                            self.sheet2[f"J{cur}"] = Rev_EQ
                            self.sheet2[f"K{cur}"] = Auth_Der
                            self.sheet2[f"L{cur}"] = B_C
                            self.sheet2[f"M{cur}"] = Rev_1
                            self.sheet2[f"N{cur}"] = Rev2

                        # increment row and write merged left/right and change reason & PD note (mirrors your code)
                        async with self.workbook_lock:
                            self.row_index += 1
                            cur2 = self.row_index
                            try:
                                self.sheet2.merge_cells(f"A{cur2}:E{cur2}")
                                self.sheet2.merge_cells(f"F{cur2}:K{cur2}")
                            except Exception:
                                pass

                        # changeReason and pdcond
                        try:
                            change_reason = await page.locator(f'[name="indirectList[{Matched_index}].changeReason"]').input_value(timeout=1000)
                        except Exception:
                            change_reason = ""
                        try:
                            PD_note = await page.locator(f'[name="indirectList[{Matched_index}].pdcond"]').input_value(timeout=1000)
                        except Exception:
                            PD_note = ""

                        async with self.workbook_lock:
                            self.sheet2[f"A{cur2}"] = change_reason
                            self.sheet2[f"F{cur2}"] = PD_note

                        async with self.workbook_lock:
                            self.row_index += 1

                        # attempt to click link text named New_JK (original used driver.find_element(By.LINK_TEXT, f'{New_JK}').click())
                        try:
                            # Playwright: find anchor with exact text and click
                            if New_JK:
                                locator = page.locator(f'a:has-text("{New_JK}")')
                                if await locator.count() > 0:
                                    try:
                                        await locator.first.click(timeout=3000)
                                    except Exception:
                                        pass
                        except Exception:
                            pass

                        entries_found += 1
                    else:
                        # not in WYS branch - still increment Matched_WY or Matched_index
                        try:
                            # attempt reading partNumber anyway
                            _ = await page.locator(f'[name="indirectList[{Matched_index}].partNumber"]').input_value(timeout=1000)
                        except Exception:
                            pass

                    Matched_WY += 1
                    Matched_index += 1

                # done processing WY section
            # end WY_exist branch

            # After full processing for this Part_no, save workbook periodically (under lock)
            async with self.workbook_lock:
                try:
                    self.workbook.save(os.path.join(self.folder_path, self.filename))
                except Exception:
                    try:
                        self.workbook.save(AUTOSAVE_PATH)
                    except Exception:
                        pass

            # close page for this worker
            try:
                await page.close()
            except Exception:
                pass

            return {"part": Part_no, "status": "OK"}
        except Exception as e:
            # ensure page closed on exceptions
            try:
                await page.close()
            except Exception:
                pass
            print("Exception in worker for part", Part_no)
            traceback.print_exc()
            return {"part": Part_no, "error": str(e)}
        finally:
            semaphore.release()

    # -------------------------
    # Public run (entry point)
    # -------------------------
    def run(self):
        # Run Tkinter mainloop briefly to allow dialogs to appear then proceed
        self.Basic()  # blocks until file selected and Part_nos loaded

        # Run the async Playwright logic
        asyncio.run(self.run_playwright_parallel())

        # Keep the tkinter mainloop until user closes (preserve earlier run behavior)
        # But we won't block forever — just show a message on completion.
        try:
            messagebox.showinfo("Done", "Processing completed. Excel saved.")
        except Exception:
            pass
        self.root.quit()


if __name__ == "__main__":
    app = App()
    app.run()
