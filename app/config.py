#!/usr/bin/env python3

import os

# ==========================
# Application constants
# ==========================

APP_W = 1024
APP_H = 640

SETUP_FILE = "setup.txt"

HH_LEN = 6          # household code length: 000629
CATEGORY_LEN = 3    # category code length: 011, 022, 098, 201, 831


# ==========================
# Setup file helpers
# ==========================

def load_setup():
    """
    Load database connection settings
    from setup.txt.

    Returns:
        dict
    """
    if not os.path.exists(SETUP_FILE):
        return {
            "host": "",
            "port": "5432",
            "database": "",
            "user": "",
            "password": "",
        }

    config = {}

    with open(
        SETUP_FILE,
        "r",
        encoding="utf-8"
    ) as file:
        for line in file:
            if "=" in line:
                key, value = line.strip().split("=", 1)
                config[key] = value

    return {
        "host": config.get("host", ""),
        "port": config.get("port", "5432"),
        "database": config.get("database", ""),
        "user": config.get("user", ""),
        "password": config.get("password", ""),
    }


def save_setup(
    host,
    port,
    database,
    user,
    password
):
    """
    Save database connection settings
    to setup.txt.
    """
    with open(
        SETUP_FILE,
        "w",
        encoding="utf-8"
    ) as file:
        file.write(f"host={host}\n")
        file.write(f"port={port}\n")
        file.write(f"database={database}\n")
        file.write(f"user={user}\n")
        file.write(f"password={password}\n")
