import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os

# Import your pipeline
try:
    from core import run_pipeline
except ImportError:
    run_pipeline = None


class BatesGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Bates Labeling & Conversion Tool")
        self.geometry("980x650")

        container = ttk.Frame(self, padding=10)
        container.pack(fill="both", expand=True)

        # ===== Form section =====
        form = ttk.Frame(container)
        form.pack(fill="x", pady=(0, 10))

        # --- Row 0: Root folder ---
        ttk.Label(form, text="Root folder:").grid(row=0, column=0, sticky="w")
        self.root_var = tk.StringVar()
        root_entry = ttk.Entry(form, textvariable=self.root_var, width=70)
        root_entry.grid(row=0, column=1, padx=5, sticky="w")
        ttk.Button(form, text="Browse", command=self.browse_folder).grid(row=0, column=2, padx=5)

        # --- Row 1: Prefix / Digits / Starting # ---
        row1_y = 1
        ttk.Label(form, text="Prefix:").grid(row=row1_y, column=0, sticky="w", pady=(8, 0))
        self.prefix_var = tk.StringVar(value="CF")
        ttk.Entry(form, textvariable=self.prefix_var, width=10).grid(
            row=row1_y, column=1, sticky="w", pady=(8, 0)
        )

        ttk.Label(form, text="Digits:").grid(
            row=row1_y, column=1, sticky="w", padx=(120, 0), pady=(8, 0)
        )
        self.digits_var = tk.StringVar(value="4")
        ttk.Entry(form, textvariable=self.digits_var, width=5).grid(
            row=row1_y, column=1, sticky="w", padx=(170, 0), pady=(8, 0)
        )

        ttk.Label(form, text="Starting #:").grid(
            row=row1_y, column=1, sticky="w", padx=(230, 0), pady=(8, 0)
        )
        self.start_var = tk.StringVar(value="1")
        ttk.Entry(form, textvariable=self.start_var, width=8).grid(
            row=row1_y, column=1, sticky="w", padx=(320, 0), pady=(8, 0)
        )

        # ===== Toggles =====

        # File-level options
        self.dry_run_var = tk.BooleanVar(value=True)
        self.backup_var = tk.BooleanVar(value=True)
        self.keep_name_var = tk.BooleanVar(value=True)

        ttk.Checkbutton(
            form,
            text="Dry run (preview only, no changes)",
            variable=self.dry_run_var,
            command=self.on_dry_run_toggle,
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        ttk.Checkbutton(
            form,
            text="Backup originals before processing (true original tree, all types)",
            variable=self.backup_var,
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(2, 0))

        ttk.Checkbutton(
            form,
            text="Append original filename after Bates (e.g. CF 0001-0008 - Original Name.pdf)",
            variable=self.keep_name_var,
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(2, 0))

        # Folder-level options
        self.rename_folders_var = tk.BooleanVar(value=False)
        self.keep_folder_name_var = tk.BooleanVar(value=True)

        self.rename_folders_cb = ttk.Checkbutton(
            form,
            text="Rename folders with Bates range (e.g. CF 0001-0244 - Folder Name)",
            variable=self.rename_folders_var,
            command=self.on_rename_folders_toggle,
        )
        self.rename_folders_cb.grid(row=5, column=0, columnspan=3, sticky="w", pady=(10, 0))

        self.keep_folder_name_cb = ttk.Checkbutton(
            form,
            text="When renaming folders, append original folder name after Bates",
            variable=self.keep_folder_name_var,
        )
        self.keep_folder_name_cb.grid(row=6, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.keep_folder_name_cb.state(["disabled"])

        # Video ordering
        self.videos_at_end_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            form,
            text="Number videos at end (after all other items)",
            variable=self.videos_at_end_var,
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(8, 0))

        # Combined final PDF
        self.combine_final_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            form,
            text="Create combined PDF for full Bates range (e.g. CF 0001- CF 0244.pdf)",
            variable=self.combine_final_var,
        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(2, 0))

        # Conversion-only mode
        self.conversion_only_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            form,
            text="Conversion-only mode (convert & format only, NO renaming or Bates)",
            variable=self.conversion_only_var,
            command=self.on_conversion_only_toggle,
        ).grid(row=9, column=0, columnspan=3, sticky="w", pady=(10, 0))

        # ===== Buttons =====
        buttons = ttk.Frame(container)
        buttons.pack(fill="x", pady=(0, 5))

        self.run_button = ttk.Button(buttons, text="Run", command=self.on_run_clicked)
        self.run_button.pack(side="left")

        ttk.Button(buttons, text="Close", command=self.destroy).pack(side="right")

        # ===== Log / Output =====
        log_frame = ttk.LabelFrame(container, text="Output")
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(log_frame, wrap="word", state="disabled")
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        # Theme
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

    # ===== UI helpers =====

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.root_var.set(folder)

    def log(self, text: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.update_idletasks()

    def clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def set_running_state(self, running: bool):
        if running:
            self.run_button.configure(text="Running...", state="disabled")
        else:
            self.run_button.configure(text="Run", state="normal")

    def on_dry_run_toggle(self):
        # In dry run, backups are harmless but also not necessary; we leave it as user choice.
        pass

    def on_rename_folders_toggle(self):
        if self.rename_folders_var.get():
            self.keep_folder_name_cb.state(["!disabled"])
        else:
            self.keep_folder_name_cb.state(["disabled"])

    def on_conversion_only_toggle(self):
        """
        Conversion-only mode ignores:
          - keep_name_var
          - folder renaming
          - combined final PDF
        So we visually gray them out when it's enabled.
        """
        conv_only = self.conversion_only_var.get()
        controls = [
            self.rename_folders_cb,
            self.keep_folder_name_cb,
        ]
        if conv_only:
            for c in controls:
                c.state(["disabled"])
            self.combine_final_var.set(False)
        else:
            self.rename_folders_cb.state(["!disabled"])
            if self.rename_folders_var.get():
                self.keep_folder_name_cb.state(["!disabled"])

    # ===== Main actions =====

    def on_run_clicked(self):
        if run_pipeline is None:
            messagebox.showerror(
                "Error",
                "Could not import run_pipeline from core.py.\n"
                "Make sure core.py is in the same folder and defines run_pipeline().",
            )
            return

        root = self.root_var.get().strip()
        prefix = self.prefix_var.get().strip()
        dry_run = self.dry_run_var.get()
        backup = self.backup_var.get()
        keep_name = self.keep_name_var.get()
        rename_folders = self.rename_folders_var.get()
        keep_folder_name = self.keep_folder_name_var.get()
        videos_at_end = self.videos_at_end_var.get()
        combine_final = self.combine_final_var.get()
        conversion_only = self.conversion_only_var.get()

        # Validate inputs
        if not root or not os.path.isdir(root):
            messagebox.showerror("Invalid folder", "Please select a valid root folder.")
            return

        if not prefix:
            messagebox.showerror("Invalid prefix", "Prefix cannot be empty.")
            return

        try:
            digits = int(self.digits_var.get())
            start = int(self.start_var.get())
        except ValueError:
            messagebox.showerror(
                "Invalid input",
                "Digits and Starting # must be whole numbers."
            )
            return

        # If conversion-only, some toggles are irrelevant
        if conversion_only and combine_final:
            messagebox.showinfo(
                "Note",
                "Combined final PDF is ignored in conversion-only mode."
            )

        self.clear_log()
        self.log(f"Root: {root}")
        self.log(f"Prefix: {prefix}")
        self.log(f"Digits: {digits}, Starting #: {start}")
        self.log(f"Dry run: {dry_run}")
        self.log(f"Backup originals: {backup}")
        self.log(f"Conversion-only mode: {conversion_only}")
        if not conversion_only:
            self.log(f"Append original filename after Bates (files): {keep_name}")
            self.log(f"Rename folders with Bates ranges: {rename_folders}")
            if rename_folders:
                self.log(f"Append original folder name after Bates: {keep_folder_name}")
            self.log(f"Number videos at end: {videos_at_end}")
            self.log(f"Create combined final PDF: {combine_final}")
        else:
            self.log("Renaming, Bates stamping, folder renaming, and combined PDF are DISABLED.")
        self.log("Starting pipeline...\n")

        self.set_running_state(True)

        # Run pipeline in a background thread so GUI stays responsive
        thread = threading.Thread(
            target=self.run_pipeline_thread,
            args=(
                root,
                prefix,
                digits,
                start,
                dry_run,
                backup,
                keep_name,
                rename_folders,
                keep_folder_name,
                videos_at_end,
                combine_final,
                conversion_only,
            ),
            daemon=True,
        )
        thread.start()

    def run_pipeline_thread(
        self,
        root,
        prefix,
        digits,
        start,
        dry_run,
        backup,
        keep_name,
        rename_folders,
        keep_folder_name,
        videos_at_end,
        combine_final,
        conversion_only,
    ):
        try:
            summary = run_pipeline(
                root_folder=root,
                prefix=prefix,
                digits=digits,
                start_counter=start,
                dry_run=dry_run,
                backup_before_bates=backup,
                keep_original_name=keep_name,
                rename_folders=rename_folders,
                keep_original_folder_name=keep_folder_name,
                number_videos_at_end=videos_at_end,
                combine_final=combine_final,
                conversion_only=conversion_only,
            )
            self.after(0, self.display_summary, summary)
        except Exception as e:
            self.after(0, self.handle_error, e)

    def display_summary(self, summary):
        if summary is None:
            self.log("\nPipeline finished with no summary data.")
            messagebox.showinfo("Done", "Pipeline finished.")
            self.set_running_state(False)
            return

        self.log("\nPipeline completed.\n")

        total_files = summary.get("total_files", "N/A")
        total_pages = summary.get("total_pages", "N/A")
        renamed = summary.get("renamed", [])
        skipped = summary.get("skipped", [])
        errors = summary.get("errors", [])

        self.log(f"Total items processed: {total_files}")
        self.log(f"Total pages (PDFs): {total_pages}\n")

        if renamed:
            self.log("Renamed / Generated items:")
            for old, new in renamed:
                self.log(f"  {old}  ->  {new}")
            self.log("")

        if skipped:
            self.log("Skipped:")
            for s in skipped:
                self.log(f"  {s}")
            self.log("")

        if errors:
            self.log("Errors / Warnings:")
            for e in errors:
                self.log(f"  {e}")
            messagebox.showwarning(
                "Completed with issues",
                "The pipeline completed, but some items had errors. See output for details.",
            )
        else:
            messagebox.showinfo("Done", "Pipeline completed successfully.")

        self.set_running_state(False)

    def handle_error(self, e: Exception):
        self.log(f"\nError: {e}")
        messagebox.showerror("Error", str(e))
        self.set_running_state(False)


def main():
    app = BatesGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
