# -*- coding: utf-8 -*-

from .libs import os, tk, open_new
from . import __version__


class About(tk.Toplevel):
    def __init__(self, version=__version__):
        super().__init__()
        self.resizable(width=False, height=False)
        if os.name == "nt" and os.path.exists("images/pyc.ico"):
            self.wm_iconbitmap("images/pyc.ico")
        self.version = version
        self.name = "WFM-BPGen"
        self.date_built = "01.05.2022"
        self.date_updated = "15.05.2022"
        self.developed_by = "Tanberk Celalettin Kutlu"
        self.contact = "tckutlu@gmail.com"
        self.github = "https://github.com/dildeolupbiten/WFM-BPGen"
        self.top_frame = tk.Frame(
            master=self,
            bd="2",
            relief="groove"
        )
        self.top_frame.pack(fill="both")
        self.bottom_frame = tk.Frame(master=self)
        self.bottom_frame.pack(fill="both")
        self.title = tk.Label(
            master=self.top_frame,
            text=self.name,
            font="Arial 25"
        )
        self.title.pack()
        for i, text in enumerate(
            [
                "Version",
                "Date Built",
                "Date Updated",
                "Developed By",
                "Contact",
                "Github"
            ]
        ):
            label = tk.Label(
                master=self.bottom_frame,
                text=text,
                font="Arial 12",
                fg="red"
            )
            label.grid(row=i, column=0, sticky="w")
            double_dot = tk.Label(
                master=self.bottom_frame,
                text=":",
                font="Arial 12",
                fg="red"
            )
            double_dot.grid(row=i, column=1, sticky="w")
        for i, j in enumerate(
            (
                self.version,
                self.date_built,
                self.date_updated,
                self.developed_by,
                self.contact,
                self.github
            )
        ):
            if j in [self.contact, self.github]:
                info = tk.Label(
                    master=self.bottom_frame,
                    text=j,
                    font="Arial 12",
                    fg="blue",
                    cursor="hand2"
                )
                if j == self.contact:
                    url = "mailto://tckutlu@gmail.com"
                else:
                    url = self.github
                info.bind(
                    "<Button-1>",
                    lambda event, link=url: open_new(link))
            else:
                info = tk.Label(
                    master=self.bottom_frame,
                    text=j,
                    font="Arial 12"
                )
            info.grid(row=i, column=2, sticky="w")
