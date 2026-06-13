#!/usr/bin/env python3

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from config import (
    APP_W,
    APP_H,
    load_setup,
    save_setup
)

from db import DatabaseManager
from reference_loader import ReferenceLoader
from processors import WeeklyDataProcessor
from excel_export import ExcelExporter


class WeeklyPerformanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Weekly Performance Analyzer")
        self.root.geometry(f"{APP_W}x{APP_H}")

        self.db = DatabaseManager()
        self.reference_loader = ReferenceLoader()
        self.processor = WeeklyDataProcessor()
        self.exporter = ExcelExporter()

        self.build_gui()

        # Load saved DB config
        setup = load_setup()

        self.host_var.set(setup["host"])
        self.port_var.set(setup["port"])
        self.db_var.set(setup["database"])
        self.user_var.set(setup["user"])
        self.pass_var.set(setup["password"])

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.exit_app
        )

    # ==========================
    # GUI
    # ==========================
    def build_gui(self):
        frm = tk.Frame(self.root)
        frm.pack(
            padx=10,
            pady=10,
            fill="x"
        )

        self.db_var = tk.StringVar()
        self.host_var = tk.StringVar()
        self.port_var = tk.StringVar(
            value="5432"
        )
        self.user_var = tk.StringVar()
        self.pass_var = tk.StringVar()

        fields = [
            ("Database", self.db_var),
            ("Host", self.host_var),
            ("Port", self.port_var),
            ("User", self.user_var),
            ("Password", self.pass_var),
        ]

        for i, (label, variable) in enumerate(fields):
            tk.Label(
                frm,
                text=label
            ).grid(
                row=i,
                column=0,
                sticky="w"
            )

            tk.Entry(
                frm,
                textvariable=variable,
                width=40,
                show="*" if label == "Password" else ""
            ).grid(
                row=i,
                column=1,
                padx=5,
                pady=2
            )

        tk.Button(
            frm,
            text="Connect",
            command=self.connect_db
        ).grid(
            row=0,
            column=2,
            rowspan=2,
            padx=10
        )

        row2 = tk.Frame(self.root)
        row2.pack(
            padx=10,
            pady=10,
            fill="x"
        )

        tk.Label(
            row2,
            text="Table:"
        ).pack(side="left")

        self.table_cb = ttk.Combobox(
            row2,
            width=40,
            state="readonly"
        )

        self.table_cb.pack(
            side="left",
            padx=5
        )

        self.table_cb.bind(
            "<<ComboboxSelected>>",
            self.load_weeks
        )

        tk.Label(
            row2,
            text="Week Start:"
        ).pack(
            side="left",
            padx=(20, 0)
        )

        self.week_cb = ttk.Combobox(
            row2,
            width=15,
            state="readonly"
        )

        self.week_cb.pack(
            side="left",
            padx=5
        )

        tk.Button(
            row2,
            text="Run & Save Excel",
            command=self.run_all
        ).pack(
            side="right",
            padx=10
        )

        tk.Button(
            row2,
            text="Exit",
            command=self.exit_app
        ).pack(
            side="right",
            padx=10
        )

        self.log = tk.Text(
            self.root,
            height=20
        )

        self.log.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )

    # ==========================
    # Logging
    # ==========================
    def write_log(self, text):
        self.log.insert(
            "end",
            text + "\n"
        )

        self.log.see("end")

        self.root.update_idletasks()

    # ==========================
    # Database actions
    # ==========================
    def connect_db(self):
        try:
            self.db.connect(
                dbname=self.db_var.get().strip(),
                host=self.host_var.get().strip(),
                port=self.port_var.get().strip(),
                user=self.user_var.get().strip(),
                password=self.pass_var.get(),
            )

            self.write_log(
                "Connected."
            )

            self.load_tables()

        except Exception as e:
            messagebox.showerror(
                "Connection error",
                str(e)
            )

    def load_tables(self):
        try:
            tables = self.db.get_valid_tables()

            self.table_cb["values"] = tables

            if tables:
                self.table_cb.current(0)

                self.load_weeks()

                self.write_log(
                    f"Loaded "
                    f"{len(tables)} "
                    f"compatible table(s)."
                )

            else:
                self.table_cb.set("")
                self.week_cb.set("")
                self.week_cb["values"] = []

                self.write_log(
                    "No compatible "
                    "tables found."
                )

        except Exception as e:
            self.write_log(
                f"Failed loading "
                f"tables: {e}"
            )

    def load_weeks(self, event=None):
        table = (
            self.table_cb
            .get()
            .strip()
        )

        if not table:
            self.week_cb.set("")
            self.week_cb["values"] = []
            return

        try:
            weeks = (
                self.db
                .get_available_weeks(
                    table
                )
            )

            self.week_cb["values"] = weeks

            if weeks:
                self.week_cb.current(
                    len(weeks) - 1
                )
            else:
                self.week_cb.set("")

        except Exception as e:
            self.write_log(
                f"Failed loading "
                f"weeks: {e}"
            )

    # ==========================
    # Main workflow
    # ==========================
    def run_all(self):
        table = (
            self.table_cb
            .get()
            .strip()
        )

        week_start = (
            self.week_cb
            .get()
            .strip()
        )

        if not table or not week_start:
            messagebox.showerror(
                "Error",
                "Select table and week"
            )
            return

        save_dir = filedialog.askdirectory(
            title=(
                "Select folder to save "
                "interviewer files"
            )
        )

        if not save_dir:
            return

        try:
            self.write_log(
                "Loading reference files..."
            )

            supervisors = (
                self.reference_loader
                .load_supervisors()
            )

            reference_df = (
                self.reference_loader
                .load_reference()
            )

            self.write_log(
                "Loading database data..."
            )

            raw_df = (
                self.db
                .load_week_data(table)
            )

            if raw_df.empty:
                raise Exception(
                    "Selected table "
                    "contains no data"
                )

            self.write_log(
                "Processing weekly data..."
            )

            result_bundle = (
                self.processor
                .process_week(
                    raw_df=raw_df,
                    supervisors=supervisors,
                    reference_df=reference_df,
                    week_start=week_start,
                    logger=self.write_log
                )
            )

            self.write_log(
                "Exporting Excel files..."
            )

            saved_count = (
                self.exporter.export(
                    result_bundle=result_bundle,
                    save_dir=save_dir,
                    week_start=week_start,
                    logger=self.write_log
                )
            )

            self.write_log(
                f"DONE: created "
                f"{saved_count} "
                f"interviewer file(s)"
            )

            messagebox.showinfo(
                "Done",
                f"Created "
                f"{saved_count} "
                f"interviewer file(s)."
            )

        except Exception as e:
            self.write_log(
                f"ERROR: {e}"
            )

            messagebox.showerror(
                "Error",
                str(e)
            )

    # ==========================
    # Exit
    # ==========================
    def exit_app(self):
        save_setup(
            host=self.host_var.get(),
            port=self.port_var.get(),
            database=self.db_var.get(),
            user=self.user_var.get(),
            password=self.pass_var.get()
        )

        try:
            self.db.close()
        except Exception:
            pass

        self.root.destroy()
