#!/usr/bin/env python3

import pandas as pd
import psycopg2
from psycopg2 import sql


class DatabaseManager:
    def __init__(self):
        self.conn = None

    # ==========================
    # Connection
    # ==========================
    def connect(
        self,
        dbname,
        host,
        port,
        user,
        password
    ):
        """
        Create PostgreSQL connection.
        Reconnect safely if connection exists.
        """
        self.close()

        self.conn = psycopg2.connect(
            dbname=dbname,
            host=host,
            port=port,
            user=user,
            password=password,
        )

    def close(self):
        """
        Close database connection safely.
        """
        if self.conn:
            try:
                self.conn.close()
            except Exception:
                pass
            finally:
                self.conn = None

    def is_connected(self):
        return self.conn is not None

    # ==========================
    # Table validation
    # ==========================
    def validate_table_structure(self, table):
        """
        Ensure selected table contains
        all required columns.
        """
        query = """
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = 'public'
            AND table_name = %s
        """

        columns = pd.read_sql(
            query,
            self.conn,
            params=[table]
        )["column_name"].tolist()

        required_columns = {
            "f0103",
            "f0105",
            "f0122",
            "prod_group",
            "date"
        }

        missing = required_columns - set(columns)

        if missing:
            raise Exception(
                f"Table '{table}' "
                f"is missing required columns: "
                f"{', '.join(sorted(missing))}"
            )

    # ==========================
    # Tables
    # ==========================
    def get_valid_tables(self):
        """
        Return compatible tables that
        contain all required columns.
        """
        query = """
            SELECT c.table_name
            FROM information_schema.columns c
            WHERE c.table_schema = 'public'
            GROUP BY c.table_name
            HAVING COUNT(DISTINCT CASE
                WHEN c.column_name IN (
                    'f0103',
                    'f0105',
                    'f0122',
                    'prod_group',
                    'date'
                )
                THEN c.column_name
            END) = 5
            ORDER BY c.table_name
        """

        df = pd.read_sql(
            query,
            self.conn
        )

        return df["table_name"].tolist()

    # ==========================
    # Weeks
    # ==========================
    def get_available_weeks(self, table):
        """
        Return available week start dates
        for selected table.
        """
        self.validate_table_structure(table)

        query = sql.SQL("""
            SELECT DISTINCT date
            FROM {table}
            WHERE date IS NOT NULL
            ORDER BY date
        """).format(
            table=sql.Identifier(table)
        )

        df = pd.read_sql(
            query.as_string(self.conn),
            self.conn
        )

        df["date"] = pd.to_datetime(
            df["date"],
            errors="coerce"
        )

        df = df.dropna(
            subset=["date"]
        ).copy()

        if df.empty:
            return []

        df["date"] = df["date"].dt.normalize()

        df["week_start"] = (
            df["date"]
            - pd.to_timedelta(
                df["date"].dt.weekday,
                unit="d"
            )
        )

        weeks = sorted(
            df["week_start"]
            .drop_duplicates()
            .dt.strftime("%Y-%m-%d")
            .tolist()
        )

        return weeks

    # ==========================
    # Raw data loading
    # ==========================
    def load_week_data(self, table):
        """
        Load source data from table.

        Returns raw dataframe before
        transformation/cleaning.
        """
        self.validate_table_structure(table)

        query = sql.SQL("""
            SELECT
                f0103 AS hh,
                prod_group,
                f0122 AS ean,
                date
            FROM {table}
        """).format(
            table=sql.Identifier(table)
        )

        df = pd.read_sql(
            query.as_string(self.conn),
            self.conn
        )

        return df
