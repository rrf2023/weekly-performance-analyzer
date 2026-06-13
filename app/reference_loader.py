#!/usr/bin/env python3

import os
import pandas as pd

from processors import WeeklyDataProcessor


class ReferenceLoader:
    def __init__(self):
        self.processor = (
            WeeklyDataProcessor()
        )

    # ==========================
    # Supervisors
    # ==========================
    def load_supervisors(self):
        """
        Load Supervisors.xlsx.

        Expected columns:
            S2 -> household
            S4 -> interviewer
            S5 -> region
        """

        path = os.path.join(
            os.getcwd(),
            "Supervisors.xlsx"
        )

        if not os.path.exists(path):
            raise Exception(
                "Supervisors.xlsx "
                "not found in "
                "program folder"
            )

        df = pd.read_excel(path)

        required_columns = {
            "S2",
            "S4",
            "S5"
        }

        missing = (
            required_columns
            - set(df.columns)
        )

        if missing:
            raise Exception(
                "Supervisors.xlsx "
                "is missing columns: "
                f"{', '.join(sorted(missing))}"
            )

        df = df.rename(columns={
            "S2":
                "hh_num",
            "S4":
                "Интервьюер",
            "S5":
                "Регион"
        })

        df = df[
            [
                "hh_num",
                "Интервьюер",
                "Регион"
            ]
        ].copy()

        df["hh_num"] = (
            df["hh_num"]
            .apply(
                self.processor
                .normalize_household
            )
        )

        df["Интервьюер"] = (
            df["Интервьюер"]
            .apply(
                self.processor
                .clean_text_value
            )
        )

        df["Регион"] = (
            df["Регион"]
            .apply(
                self.processor
                .clean_text_value
            )
        )

        df = df[
            df["hh_num"] != ""
        ]

        df = df.drop_duplicates(
            subset=["hh_num"]
        )

        return df

    # ==========================
    # Reference categories
    # ==========================
    def load_reference(self):
        """
        Load Справочник.xlsx.

        Expected columns:
            Категория
            Описание
        """

        path = os.path.join(
            os.getcwd(),
            "Справочник.xlsx"
        )

        if not os.path.exists(path):
            raise Exception(
                "Справочник.xlsx "
                "not found in "
                "program folder"
            )

        df = pd.read_excel(path)

        df.columns = [
            str(col).strip()
            for col in df.columns
        ]

        required_columns = {
            "Категория",
            "Описание"
        }

        missing = (
            required_columns
            - set(df.columns)
        )

        if missing:
            raise Exception(
                "Справочник.xlsx "
                "is missing columns: "
                f"{', '.join(sorted(missing))}"
            )

        df = df[
            [
                "Категория",
                "Описание"
            ]
        ].copy()

        df["Категория"] = (
            df["Категория"]
            .apply(
                self.processor
                .normalize_category
            )
        )

        df["Описание"] = (
            df["Описание"]
            .apply(
                self.processor
                .clean_text_value
            )
        )

        df = df[
            df["Категория"] != ""
        ]

        df = df.drop_duplicates(
            subset=["Категория"]
        )

        df = df.rename(columns={
            "Категория":
                "ref_category"
        })

        return df
