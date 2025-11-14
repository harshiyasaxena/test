def show_options_dialog(self):
    dlg = tk.Toplevel(self.root)
    dlg.title("Options")
    dlg.grab_set()

    # state holders
    self.fig_file = None
    self.epl_folder = None
    self.future1_folder = None
    self.future2_folder = None

    # BUTTON FUNCTIONS
    def pick_fig():
        path = filedialog.askopenfilename(
            title="Select Figure PDF", filetypes=[("PDF files", "*.pdf")]
        )
        if path:
            self.fig_file = path
            btn1.config(text=f"FIG Selected")

    def pick_epl():
        path = filedialog.askdirectory(title="Select EPL Folder")
        if path:
            self.epl_folder = path
            btn2.config(text="EPL Selected")

    def pick_future1():
        path = filedialog.askdirectory(title="Select Folder")
        if path:
            self.future1_folder = path
            btn3.config(text="Future1 Selected")

    def pick_future2():
        path = filedialog.askdirectory(title="Select Folder")
        if path:
            self.future2_folder = path
            btn4.config(text="Future2 Selected")

    # UI layout
    frm = tk.Frame(dlg)
    frm.pack(padx=10, pady=10)

    btn1 = tk.Button(frm, text="Import Figure Content", width=25, command=pick_fig)
    btn1.grid(row=0, column=0, padx=5, pady=5)

    btn2 = tk.Button(frm, text="Import EPL", width=25, command=pick_epl)
    btn2.grid(row=1, column=0, padx=5, pady=5)

    btn3 = tk.Button(frm, text="Future", width=25, command=pick_future1)
    btn3.grid(row=2, column=0, padx=5, pady=5)

    btn4 = tk.Button(frm, text="Future", width=25, command=pick_future2)
    btn4.grid(row=3, column=0, padx=5, pady=5)

    # RUN BUTTON
    def on_run():
        dlg.grab_release()
        dlg.destroy()

        self.start_timer()

        # ---- Run FIGURE ----
        if self.fig_file:
            try:
                out = os.path.splitext(self.fig_file)[0] + ".xlsx"
                pdf_to_excel(self.fig_file, out)
                self.excel_file = out
                self.folder_path = os.path.dirname(out)
                self.filename = os.path.basename(out)

                if hasattr(self, "Basic"):
                    self.Basic()
                if hasattr(self, "main_code"):
                    self.main_code()

            except Exception as e:
                messagebox.showerror("Figure Error", f"{e}")

        # ---- Run EPL ----
        if self.epl_folder:
            try:
                self.folder_path = self.epl_folder
                self.browse_and_EPL()   # already does the job
            except Exception as e:
                messagebox.showerror("EPL Error", f"{e}")

        # ---- Future actions (do nothing for now) ----
        if self.future1_folder:
            print("Future Task 1 placeholder")

        if self.future2_folder:
            print("Future Task 2 placeholder")

        self.exit_app()

    run_btn = tk.Button(frm, text="RUN", width=15, bg="green", fg="white", command=on_run)
    run_btn.grid(row=0, column=1, rowspan=4, padx=15)

    # CANCEL BUTTON
    def on_cancel():
        dlg.grab_release()
        dlg.destroy()
        self.exit_app()

    cancel_btn = tk.Button(dlg, text="Cancel", command=on_cancel)
    cancel_btn.pack(pady=5)
