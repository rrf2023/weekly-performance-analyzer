#!/usr/bin/env python3

import re
from datetime import timedelta

import pandas as pd

from config import HH_LEN, CATEGORY_LEN


class WeeklyDataProcessor:
    # ==========================
    # Helpers / normalization
    # ==========================
    @staticmethod
    def clean_text_value(value):
        if pd.isna(value):
            return ""
        return str(value).strip()

    @staticmethod
    def normalize_household(value):
        """
        Convert household code to zero-padded 6-digit string.

        Examples:
            629 -> 000629
            '1174.0' -> 001174
            '000629' -> 000629
        """
        if pd.isna(value):
            return ""

        text = str(value).strip()

        if text.endswith(".0"):
            text = text[:-2]

        digits = re.sub(r"\D", "", text)

        if not digits:
            return ""

        return digits.zfill(HH_LEN)

    @staticmethod
    def normalize_category(value):
        """
        Convert category code to zero-padded
        3-digit string.
        """
        if pd.isna(value):
            return ""

        text = str(value).strip()

        if text.endswith(".0"):
            text = text[:-2]

        digits = re.sub(r"\D", "", text)

        if not digits:
            return ""

        return digits.zfill(CATEGORY_LEN)

    @staticmethod
    def normalize_ean(value):
        """
        Convert EAN/SKU to digit string
        without losing leading zeroes.
        """
        if pd.isna(value):
            return ""

        text = str(value).strip()

        if text.endswith(".0"):
            text = text[:-2]

        return re.sub(r"\D", "", text)

    @staticmethod
    def safe_filename(text):
        text = str(text).strip()

        if not text:
            text = "Без_интервьюера"

        text = re.sub(r'[\\/:*?"<>|]+', "_", text)
        text = re.sub(r"\s+", "_", text)

        return text[:120]

    # ==========================
    # Suspicious HH logic
    # ==========================
    @staticmethod
    def build_household_signature_map(
        category_part,
        sku_part
    ):
        """
        Build suspicious household summary.

        Rules:
        - Flag 1:
          1 or 2 categories only
        - Flag 2:
          duplicated receipt
          (identical EAN composition)
        """

        if category_part.empty:
            return pd.DataFrame(columns=[
                "Домохозяйство",
                "Интервьюер",
                "Регион",
                "Кол-во категорий",
                "Признак_1_1или2_категории",
                "Признак_2_копия_чека",
                "Сигнатура"
            ])

        household_rows = []
        ean_signature_to_hhs = {}

        category_part = category_part.copy()

        category_part["Категория"] = (
            category_part["Категория"]
            .astype(str)
            .str.strip()
        )

        category_part["Количество покупок"] = pd.to_numeric(
            category_part["Количество покупок"],
            errors="coerce"
        ).fillna(0).astype(int)

        sku_part = sku_part.copy()

        if not sku_part.empty:
            sku_part["Домохозяйство"] = (
                sku_part["Домохозяйство"]
                .astype(str)
                .str.strip()
            )

            sku_part["EAN"] = (
                sku_part["EAN"]
                .astype(str)
                .str.strip()
            )

            sku_part["Количество покупок"] = pd.to_numeric(
                sku_part["Количество покупок"],
                errors="coerce"
            ).fillna(0).astype(int)

        for hh, hh_part in category_part.groupby(
            "Домохозяйство",
            dropna=False
        ):
            hh_part = hh_part.copy()
            hh_str = str(hh).strip()

            interviewer = ""
            region = ""

            if (
                "Интервьюер" in hh_part.columns
                and not hh_part.empty
            ):
                interviewer = str(
                    hh_part["Интервьюер"].iloc[0]
                ).strip()

            if (
                "Регион" in hh_part.columns
                and not hh_part.empty
            ):
                region = str(
                    hh_part["Регион"].iloc[0]
                ).strip()

            unique_categories = sorted(
                set(hh_part["Категория"].tolist())
            )

            category_count = len(
                unique_categories
            )

            category_agg = (
                hh_part
                .groupby(
                    "Категория",
                    dropna=False
                )["Количество покупок"]
                .sum()
                .reset_index()
                .sort_values(
                    by="Категория",
                    ascending=True
                )
            )

            category_signature_items = [
                f"{row['Категория']}:{int(row['Количество покупок'])}"
                for _, row
                in category_agg.iterrows()
                if str(
                    row["Категория"]
                ).strip()
            ]

            category_signature = " | ".join(
                category_signature_items
            )

            sku_hh = sku_part[
                sku_part["Домохозяйство"]
                == hh_str
            ].copy()

            ean_signature = ""

            if not sku_hh.empty:
                ean_agg = (
                    sku_hh
                    .groupby(
                        "EAN",
                        dropna=False
                    )["Количество покупок"]
                    .sum()
                    .reset_index()
                    .sort_values(
                        by="EAN",
                        ascending=True
                    )
                )

                ean_signature_items = [
                    f"{row['EAN']}:"
                    f"{int(row['Количество покупок'])}"
                    for _, row
                    in ean_agg.iterrows()
                    if str(
                        row["EAN"]
                    ).strip()
                ]

                ean_signature = " | ".join(
                    ean_signature_items
                )

                if ean_signature.strip():
                    ean_signature_to_hhs.setdefault(
                        ean_signature,
                        set()
                    ).add(hh_str)

            household_rows.append({
                "Домохозяйство":
                    hh_str,
                "Интервьюер":
                    interviewer,
                "Регион":
                    region,
                "Кол-во категорий":
                    category_count,
                "Признак_1_1или2_категории":
                    "Да"
                    if category_count in (1, 2)
                    else "",
                "Признак_2_копия_чека":
                    "",
                "Сигнатура":
                    category_signature
            })

        duplicated_households = set()

        for _, hhs in (
            ean_signature_to_hhs.items()
        ):
            if len(hhs) >= 2:
                duplicated_households.update(
                    hhs
                )

        for row in household_rows:
            if (
                row["Домохозяйство"]
                in duplicated_households
            ):
                row[
                    "Признак_2_копия_чека"
                ] = "Да"

        output = pd.DataFrame(
            household_rows
        )

        output = output[
            (
                output[
                    "Признак_1_1или2_категории"
                ] == "Да"
            )
            |
            (
                output[
                    "Признак_2_копия_чека"
                ] == "Да"
            )
        ].copy()

        return output.sort_values(
            by=[
                "Интервьюер",
                "Регион",
                "Домохозяйство"
            ]
        )

    # ==========================
    # Main processing
    # ==========================
    def process_week(
        self,
        raw_df,
        supervisors,
        reference_df,
        week_start,
        logger=None
    ):
        """
        Main weekly processing logic.
        """

        raw_df = raw_df.copy()

        raw_df["hh"] = raw_df["hh"].apply(
            self.normalize_household
        )

        raw_df["prod_group"] = (
            raw_df["prod_group"]
            .apply(
                self.normalize_category
            )
        )

        raw_df["ean"] = (
            raw_df["ean"]
            .apply(self.normalize_ean)
        )

        raw_df["date"] = pd.to_datetime(
            raw_df["date"],
            errors="coerce"
        )

        raw_df = raw_df[
            (raw_df["hh"] != "")
            &
            (raw_df["prod_group"] != "")
            &
            (
                raw_df["date"]
                .notna()
            )
        ].copy()

        if raw_df.empty:
            raise Exception(
                "No valid rows found "
                "after cleaning source data"
            )

        raw_df["date"] = (
            raw_df["date"]
            .dt.normalize()
        )

        week_start_dt = pd.to_datetime(
            week_start
        ).normalize()

        week_end = (
            week_start_dt
            + timedelta(days=6)
        )

        df_week = raw_df[
            (
                raw_df["date"]
                >= week_start_dt
            )
            &
            (
                raw_df["date"]
                <= week_end
            )
        ].copy()

        if df_week.empty:
            raise Exception(
                f"No data found "
                f"for selected week: "
                f"{week_start_dt:%Y-%m-%d}"
            )

        if logger:
            logger(
                f"Rows in selected week: "
                f"{len(df_week)}"
            )

        result = (
            df_week
            .groupby(
                ["hh", "prod_group"],
                dropna=False
            )
            .size()
            .reset_index(
                name="Количество покупок"
            )
        )

        sku_result = (
            df_week[
                df_week["ean"] != ""
            ]
            .groupby(
                ["hh", "ean"],
                dropna=False
            )
            .size()
            .reset_index(
                name="Количество покупок"
            )
        )

        result = result.merge(
            supervisors,
            left_on="hh",
            right_on="hh_num",
            how="left"
        )

        result = result.merge(
            reference_df,
            left_on="prod_group",
            right_on="ref_category",
            how="left"
        )

        result = result.drop(
            columns=[
                c for c in [
                    "hh_num",
                    "ref_category"
                ]
                if c in result.columns
            ]
        )

        result = result.rename(columns={
            "hh":
                "Домохозяйство",
            "prod_group":
                "Категория"
        })

        sku_result = sku_result.merge(
            supervisors,
            left_on="hh",
            right_on="hh_num",
            how="left"
        )

        sku_result = sku_result.drop(
            columns=[
                c for c in [
                    "hh_num"
                ]
                if c
                in sku_result.columns
            ]
        )

        sku_result = sku_result.rename(columns={
            "hh":
                "Домохозяйство",
            "ean":
                "EAN"
        })

        suspicious_parts = []

        for interviewer, part in result.groupby(
            "Интервьюер",
            dropna=False
        ):
            sku_part = sku_result[
                sku_result[
                    "Интервьюер"
                ]
                .fillna("")
                .astype(str)
                .str.strip()
                == str(interviewer).strip()
            ]

            suspicious = (
                self
                .build_household_signature_map(
                    part,
                    sku_part
                )
            )

            if not suspicious.empty:
                suspicious_parts.append(
                    suspicious
                )

        suspicious_all = (
            pd.concat(
                suspicious_parts,
                ignore_index=True
            )
            if suspicious_parts
            else pd.DataFrame()
        )

        return {
            "result":
                result,
            "sku_result":
                sku_result,
            "suspicious":
                suspicious_all,
            "week_start":
                week_start_dt
        }
