#!/usr/bin/env python3

import tkinter as tk

from gui import WeeklyPerformanceApp


def main():
    root = tk.Tk()
    app = WeeklyPerformanceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
