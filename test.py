import os
import time
import string
import asyncio
import traceback
from typing import Optional, List

import tkinter as tk
from tkinter import filedialog, messagebox

import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError, Page, Browser

# ---------- Configuration ----------
BASE_URL = "https://hihi.com"             # your real site
CONCURRENCY = 4                            # number of parallel pages/tasks
PAGE_TIMEOUT = 20000                       # ms timeout for Playwright waits
HEADLESS = True                            # set False to watch the browser open
# -----------------------------------


class App:
    def __init__(self):
        # tkinter kept for compatibility — we will not show UI by default (same as original)
        self.root = tk.Tk()
        self.root.title("PULSE")
        self.root.withdraw()

        # Excel / workbook related
        self.Part_nos: List[str] = []
        self.folder_path: Optional[str] = None
        self.filename: Optional[str] = None
        self.workbook = None
        self.sheet1 = None
        self.sheet2 = None
        self.sheet3 = None
        self.row_index = 1

        # Formatting
        self.green = PatternFill(start_color='c6efce', end_color='c6efce', fill_type='solid')
        self.red = PatternFill(start_color='ffc7ce', end_color='ffc7ce', fill_type='solid')
        thin = Side(style='thin')
        self.Thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)
        self.blue = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')  # choose a blue
        self.bold_font = Font(bold=True)

        # Playwright runtime objects (set in run_playwright)
        self.browser: Optional[Browser] = None
        self.context = None
        self.base_page: Optional[Page] = None

    # ---------- Excel helpers ----------
    def open_workbook(self, folder_path: str, filename: str):
        """Load workbook and prepare sheets — mirrors original behavior with safe checks."""
        self.folder_path = folder_path
        self.filename = filename
        file_path = os.path.join(folder_path, filename)
        self.workbook = load_workbook(file_path)

        # If "PULSE DATA" exists, remove and recreate
        if "PULSE DATA" in self.workbook.sheetnames:
            std = self.workbook["PULSE DATA"]
            self.workbook.remove(std)
            self.workbook.create_sheet("PULSE DATA")

        # Ensure sheets exist (INPUTS and Report Data assumed to exist)
        if "INPUTS" not in self.workbook.sheetnames:
            raise ValueError("INPUTS sheet not found in workbook.")
        if "Report Data" not in self.workbook.sheetnames:
            # create if missing
            self.workbook.create_sheet("Report Data")

        self.sheet1 = self.workbook["INPUTS"]
        self.sheet2 = self.workbook["PULSE DATA"]
        self.sheet3 = self.workbook["Report Data"]

        # Set column widths (same as original)
        col_widths = {
            'A': 20, 'B': 20, 'C': 20, 'D': 20, 'E': 10, 'F': 20,
            'G': 10, 'H': 10, 'I': 20, 'J': 20, 'K': 26, 'L': 10,
            'M': 20, 'N': 20
        }
        for col, w in col_widths.items():
            self.sheet2.column_dimensions[col].width = w

    def read_part_numbers(self):
        """Read part numbers from column B starting at row 2 (index 1 in openpyxl)."""
        self.Part_nos = []
        # openpyxl uses 1-based indexing; sheet['B'] returns cells in column B including header
        col_b = self.sheet1['B']
        # skip header (index 0)
        for cell in col_b[1:]:
            if cell.value is None:
                break
            self.Part_nos.append(str(cell.value).strip())

    # ---------- Playwright interactions ----------
    async def run_playwright(self):
        """Main orchestration: launch browser, perform initial setup, then parallel extraction."""
        async with async_playwright() as pw:
            browser = await pw.firefox.launch(headless=HEADLESS)
            self.browser = browser
            # create shared context (will be used to spawn pages for parallel tasks)
            context = await browser.new_context()
            self.context = context

            # Create initial page to perform initial clicks/flows (Continue -> new window)
            page = await context.new_page()
            self.base_page = page
            page.set_default_timeout(PAGE_TIMEOUT)

            await page.goto(BASE_URL)
            # Try to click a Continue input if present
            try:
                # try to click input with value Continue (like original)
                await page.locator('input[value="Continue"]').click(timeout=5000)
            except PlaywrightTimeoutError:
                # fallback: try to find any input and click first with value attribute Continue
                try:
                    inputs = page.locator("input")
                    count = await inputs.count()
                    for i in range(count):
                        val = await inputs.nth(i).get_attribute("value")
                        if val and val.strip() == "Continue":
                            await inputs.nth(i).click()
                            break
                except Exception:
                    pass
            except Exception:
                pass

            # Wait for potential new page (the original Selenium switched to last window)
            # If the click opens a new window/tab, Playwright can detect it.
            # Give it a short window to open.
            await asyncio.sleep(1)
            pages = context.pages
            if len(pages) > 1:
                # use last opened page as base
                self.base_page = pages[-1]

            # Now export storage state if we want, but since we're reusing same context and pages,
            # we will spawn multiple pages from the same context to reuse cookies/session.
            # Read part numbers
            self.read_part_numbers()

            # Prepare Excel starting row index
            # We'll place results starting at row 1 as in original; however preserve last used row if needed
            self.row_index = 1

            # Use a semaphore to limit concurrency
            sem = asyncio.Semaphore(CONCURRENCY)
            tasks = []
            for part_no in self.Part_nos:
                tasks.append(self._bounded_extract(sem, part_no))

            # run all tasks
            await asyncio.gather(*tasks)

            # After all extraction done, apply alignment and save workbook
            self._finalize_and_save()

            # close context and browser
            await context.close()
            await browser.close()

    async def _bounded_extract(self, sem: asyncio.Semaphore, part_no: str):
        """Wrap each extraction task with semaphore to limit concurrency."""
        async with sem:
            try:
                await self.extract_for_part(part_no)
            except Exception:
                traceback.print_exc()

    async def extract_for_part(self, Part_no: str):
        """Translate your main_code logic for one Part_no into Playwright operations and write to Excel."""
        page = await self.context.new_page()
        page.set_default_timeout(PAGE_TIMEOUT)

        # Save workbook frequently (original saved each loop)
        try:
            self.workbook.save(os.path.join(self.folder_path, self.filename))
        except Exception:
            pass

        # Example initial writes (mirrors original)
        # Note: openpyxl append will place values at next free row by design; we replicate more control
        # We'll write cells explicitly using current self.row_index and increment accordingly,
        # but because tasks run in parallel we must protect writes to workbook (not thread-safe).
        # So we implement a simple lock around workbook writes using asyncio.Lock.
        # For simplicity and safety, we'll collect rows in-memory and append them using a write lock.

        # We will use an asyncio.Lock attached to self for workbook writes
        if not hasattr(self, "_wb_lock"):
            self._wb_lock = asyncio.Lock()

        # 1) Initialize header-ish rows similar to original first writes
        async with self._wb_lock:
            r = self.row_index
            self.sheet2[f'A{r}'] = "JK"
            self.sheet2[f'B{r}'] = "Def Mef"
            self.sheet2[f'C{r}'] = "JIR"

            # color first three columns
            for let in string.ascii_uppercase[:3]:
                self.sheet2[f'{let}{r}'].fill = self.blue
                self.sheet2[f'{let}{r}'].font = Font(color="ffffff", bold=True)
                self.sheet2[f'{let}{r}'].border = self.Thin_border

            self.row_index += 1
            current_row_for_part = self.row_index
            self.sheet2[f'A{current_row_for_part}'] = Part_no
            self.row_index += 1

            # Save index snapshot for writing later values to avoid race on self.row_index
            start_row = current_row_for_part

        # Fill the search box and trigger JS as in original
        try:
            # find search element
            try:
                await page.fill("#searchkalu", Part_no, timeout=5000)
            except PlaywrightTimeoutError:
                # sometimes element id contains periods; try alternative find
                try:
                    el = page.locator('[id="searchkalu"]')
                    await el.fill(Part_no)
                except Exception:
                    pass
            except Exception:
                pass

            # Execute original JS call
            try:
                await page.evaluate("showFile('1');")
            except Exception:
                pass

            # small wait to allow page to update
            await asyncio.sleep(0.5)

            # get def_mef and jir values from DOM
            def_mef = None
            jir = None
            try:
                def_mef = await page.get_attribute('#moderbean\\.defmed', 'value')
            except Exception:
                # try alternate id with dot as part of id (escaped)
                try:
                    def_mef = await page.locator('[id="moderbean.defmed"]').get_attribute('value')
                except Exception:
                    def_mef = ""

            try:
                jir = await page.get_attribute('#moderbean\\.jir', 'value')
            except Exception:
                try:
                    jir = await page.locator('[id="moderbean.jir"]').get_attribute('value')
                except Exception:
                    jir = ""

            # Write def_mef and jir back into Excel (acquire lock)
            async with self._wb_lock:
                self.sheet2[f'B{start_row}'] = def_mef
                self.sheet2[f'C{start_row}'] = jir

                # color these cells same as earlier
                for let in string.ascii_uppercase[:3]:
                    self.sheet2[f'{let}{start_row}'].fill = self.blue
                    self.sheet2[f'{let}{start_row}'].border = self.Thin_border

            # Next: write FC-related rows similar to original. We'll attempt to replicate flows:
            # simulate clicking showFile('4') if needed
            try:
                await page.evaluate("showFile('4');")
            except Exception:
                pass

            await asyncio.sleep(0.4)

            # FC codes that we consider interesting
            FC_text_code = ['1', '2', '3', '4', 'MF', 'AB']
            # Locate all inputs starting with name fcthoda[
            fc_inputs = page.locator('//input[starts-with(@name,"fcthoda[")]')
            count_fc_inputs = await fc_inputs.count()
            Matched_index = 0

            # We'll iterate by index because your later code used Matched_index
            for idx in range(count_fc_inputs):
                try:
                    # exact_fc element by name 'fcbean[{Matched_index}].fc' (note case variations in original)
                    name_selector = f'input[name="fcbean[{Matched_index}].fc"], input[name="fcBean[{Matched_index}].fc"]'
                    exact = page.locator(name_selector)
                    if await exact.count() == 0:
                        Matched_index += 1
                        continue
                    FC_code = await exact.first.get_attribute('value')
                except Exception:
                    Matched_index += 1
                    continue

                if FC_code in FC_text_code:
                    # Acquire write lock to safely write rows
                    async with self._wb_lock:
                        r = self.row_index
                        self.sheet2[f'A{r}'] = "FC Code"
                        self.sheet2[f'A{r}'].fill = self.blue
                        self.sheet2[f'A{r}'].border = self.Thin_border
                        self.sheet2[f'A{r}'].font = Font(color="ffffff", bold=True)

                        merge_title = f'B{r}:C{r}'
                        fc_title = ""
                        try:
                            fc_title = await page.get_attribute(f'input[name="fcBean[{Matched_index}].fcTitle"]', 'value')
                        except Exception:
                            try:
                                fc_title = await page.get_attribute(f'input[name="fcbean[{Matched_index}].fcTitle"]', 'value')
                            except Exception:
                                fc_title = ""

                        self.sheet2[f'B{r}'] = fc_title
                        # merging cells
                        try:
                            self.sheet2.merge_cells(merge_title)
                        except Exception:
                            pass

                        # FC text
                        frm = ""
                        try:
                            frm = await page.get_attribute(f'input[name="FCText[{Matched_index}].freeFormTxt"]', 'value')
                        except Exception:
                            try:
                                frm = await page.get_attribute(f'input[name="fcText[{Matched_index}].overLengthPArt"]', 'value')
                            except Exception:
                                try:
                                    frm = await page.get_attribute(f'input[name="fcText[{Matched_index}].specNum"]', 'value')
                                except Exception:
                                    frm = ""

                        self.sheet2[f'D{r}'] = frm
                        try:
                            self.sheet2.merge_cells(f'D{r}:L{r}')
                        except Exception:
                            pass

                        rev_by = ""
                        try:
                            rev_by = await page.get_attribute(f'input[name="fcText[{Matched_index}].revisedDatestr"]', 'value')
                        except Exception:
                            try:
                                rev_by = await page.get_attribute(f'input[name="fcText[{Matched_index}].revisedDatestr"]', 'value')
                            except Exception:
                                rev_by = ""

                        self.sheet2[f'M{r}'] = rev_by
                        self.sheet2[f'N{r}'] = rev_by

                        self.row_index += 1

                Matched_index += 1

            # Next: WY image check (original had WY_exist list)
            try:
                WY_exist = ["https://hehe.png", "https://hihi.png"]
                img_src = ""
                try:
                    img_src = await page.get_attribute('#iw_img', 'src')
                except Exception:
                    img_src = ""
                await asyncio.sleep(0.2)

                if img_src in WY_exist:
                    # click indirect button
                    try:
                        await page.click('button[name="bt_INDIRECTWY"]', timeout=3000)
                    except Exception:
                        # try other selectors
                        try:
                            await page.click('[name="bt_INDIRECTWY"]')
                        except Exception:
                            pass
                    await asyncio.sleep(0.4)

                    async with self._wb_lock:
                        r = self.row_index
                        headers = ["Dir/Indir", "Indir Active", "JK", "Type", "MN/CN", "MN/CN JK",
                                   "WY", "AW", "Okay used on", "Rev", "Derived", "B or C", "Revised"]
                        for idx_h, val in enumerate(headers):
                            col = string.ascii_uppercase[idx_h]
                            self.sheet2[f'{col}{r}'] = val
                            self.sheet2[f'{col}{r}'].fill = self.blue
                            self.sheet2[f'{col}{r}'].font = Font(color="ffffff", bold=True)
                            self.sheet2[f'{col}{r}'].border = self.Thin_border

                        # merge M:N
                        try:
                            self.sheet2.merge_cells(f'M{r}:N{r}')
                        except Exception:
                            pass

                        self.row_index += 1

                    # iterate WY list items (original used IndirectWYlist)
                    Matched_WY = 0
                    while True:
                        try:
                            # find WY id value
                            wy_selector = f'input[name="IndirectWYlist[{Matched_WY}].wycode"]'
                            wy_el = page.locator(wy_selector)
                            if await wy_el.count() == 0:
                                # maybe pagination or no more; try older next page button
                                try:
                                    next_page_btn = page.locator('[name="bt_Page"]')
                                    if await next_page_btn.is_enabled():
                                        await next_page_btn.click()
                                        Matched_WY = 0
                                        await asyncio.sleep(0.4)
                                        continue
                                    else:
                                        break
                                except Exception:
                                    break

                            WY_id = await wy_el.first.get_attribute('value')
                        except PlaywrightTimeoutError:
                            break
                        except Exception:
                            break

                        # WYS codes considered
                        WYS = ['1', '2', '5']
                        if str(WY_id) in WYS:
                            # collect many fields and write to sheet row
                            async with self._wb_lock:
                                r = self.row_index
                                # Direct/Indirect
                                try:
                                    direct_val = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].direct"]', 'value')
                                except Exception:
                                    direct_val = ""
                                if direct_val == "true":
                                    self.sheet2[f'A{r}'] = "Direct"
                                else:
                                    self.sheet2[f'A{r}'] = "Indirect"

                                try:
                                    status_val = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].status"]', 'value')
                                    if status_val in ["A", "Active"]:
                                        self.sheet2[f'B{r}'] = "Active"
                                        # bold
                                        self.sheet2[f'B{r}'].font = self.bold_font
                                except Exception:
                                    pass

                                try:
                                    new_jk = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].partNumber"]', 'value')
                                except Exception:
                                    new_jk = ""
                                self.sheet2[f'C{r}'] = new_jk
                                self.sheet2[f'C{r}'].fill = self.green

                                try:
                                    type_name = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].typeName"]', 'value')
                                except Exception:
                                    type_name = ""
                                self.sheet2[f'D{r}'] = type_name

                                try:
                                    mncn = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].mncn"]', 'value')
                                    self.sheet2[f'E{r}'] = mncn
                                except Exception:
                                    pass

                                try:
                                    old_jk = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].oldiePArt"]', 'value')
                                except Exception:
                                    old_jk = ""
                                self.sheet2[f'F{r}'] = old_jk
                                # fill red if needed
                                self.sheet2[f'F{r}'].fill = self.red

                                try:
                                    wy_type = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].wycode"]', 'value')
                                except Exception:
                                    wy_type = ""
                                self.sheet2[f'G{r}'] = wy_type

                                try:
                                    rework = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].rework"]', 'value')
                                except Exception:
                                    rework = ""
                                self.sheet2[f'H{r}'] = rework

                                try:
                                    applicable = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].app"]', 'value')
                                except Exception:
                                    applicable = ""
                                self.sheet2[f'I{r}'] = applicable

                                try:
                                    rev_eq = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].eq"]', 'value')
                                except Exception:
                                    rev_eq = ""
                                self.sheet2[f'J{r}'] = rev_eq

                                try:
                                    auth_der = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].authe"]', 'value')
                                except Exception:
                                    auth_der = ""
                                self.sheet2[f'K{r}'] = auth_der

                                try:
                                    b_c = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].bc"]', 'value')
                                except Exception:
                                    b_c = ""
                                self.sheet2[f'L{r}'] = b_c

                                try:
                                    rev1 = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].rev1"]', 'value')
                                except Exception:
                                    rev1 = ""
                                self.sheet2[f'M{r}'] = rev1

                                try:
                                    rev2 = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].rev2"]', 'value')
                                except Exception:
                                    rev2 = ""
                                self.sheet2[f'N{r}'] = rev2

                                self.row_index += 1

                                # Merge some cells and add extra note row as original did
                                merge_left = f'A{self.row_index}:E{self.row_index}'
                                merge_right = f'F{self.row_index}:K{self.row_index}'
                                try:
                                    self.sheet2.merge_cells(merge_left)
                                    self.sheet2.merge_cells(merge_right)
                                except Exception:
                                    pass

                                try:
                                    change_reason = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].changeReason"]', 'value')
                                except Exception:
                                    change_reason = ""
                                try:
                                    self.sheet2[f'A{self.row_index}'] = change_reason
                                except Exception:
                                    pass

                                try:
                                    pd_note = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].pdcond"]', 'value')
                                except Exception:
                                    pd_note = ""
                                try:
                                    self.sheet2[f'F{self.row_index}'] = pd_note
                                except Exception:
                                    pass

                                self.row_index += 1

                        else:
                            # Not in WYS, still attempt to read partNumber to keep New_JK variable for link-click fallback
                            try:
                                new_jk = await page.get_attribute(f'input[name="indirectList[{Matched_WY}].partNumber"]', 'value')
                            except Exception:
                                new_jk = ""
                        Matched_WY += 1

                    # After WY loop, try clicking the redirect link for New_JK if present
                    try:
                        if new_jk:
                            # try link text click
                            await page.click(f'a:has-text("{new_jk}")', timeout=2000)
                    except Exception:
                        # fallback attempts similar to original
                        try:
                            await page.click('#Filed', timeout=2000)
                        except Exception:
                            pass

            # close the page created for this part
            await page.close()

        except Exception:
            await page.close()
            traceback.print_exc()

    def _finalize_and_save(self):
        # Center align all cells and wrap text as original
        try:
            for row in self.sheet2.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        except Exception:
            pass

        # Save workbook and open file (like original)
        out_path = os.path.join(self.folder_path, self.filename)
        self.workbook.save(out_path)
        try:
            os.startfile(out_path)
        except Exception:
            # fallback for platforms without os.startfile
            print("Saved workbook to:", out_path)

    # ---------- Public run flow ----------
    def run(self):
        # Ask user for Excel file (since original had a GUI file selection flow)
        # For compatibility, we will ask user via file dialog if folder/filename not set
        if not self.folder_path or not self.filename:
            # Ask user to select file
            file = filedialog.askopenfilename(title="Select PULSE Excel file", filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xltx;*.xltm")])
            if not file:
                messagebox.showerror("No file", "No Excel file selected. Exiting.")
                return
            folder, fname = os.path.split(file)
            self.open_workbook(folder, fname)
        else:
            self.open_workbook(self.folder_path, self.filename)

        # Run Playwright tasks (async) and block until done
        asyncio.run(self.run_playwright())

        # Keep tkinter mainloop if you need the app UI (original had root.mainloop())
        # We'll exit without showing GUI; if you want the GUI active, uncomment below:
        # self.root.deiconify()
        # self.root.mainloop()


if __name__ == "__main__":
    app = App()
    app.run()
