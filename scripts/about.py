# -*- coding: utf-8 -*-

from . import __version__
from .libs import (
    os, tk, json, open_new, urlopen, Popen,
    URLError, showwarning, showinfo
)


def check_update():
    try:
        new = urlopen(
            "https://raw.githubusercontent.com/dildeolupbiten"
            "/WFM-BPGen/master/README.md"
        ).read().decode()
    except URLError:
        showwarning(
            title="Warning",
            message="Couldn't connect.",
        )
        return
    with open("README.md", "r", encoding="utf-8") as f:
        old = f.read()[:-1]
    if new != old:
        with open("README.md", "w", encoding="utf-8") as f:
            f.write(new)
    try:
        scripts = json.load(
            urlopen(
                url=f"https://api.github.com/repos/dildeolupbiten/"
                    f"WFM-BPGen/contents/scripts?ref=master"
            )
        )
    except URLError:
        showwarning(
            title="Warning",
            message="Couldn't connect."
        )
        return
    update = False
    for i in scripts:
        try:
            file = urlopen(i["download_url"]).read().decode()
        except URLError:
            showwarning(
                title="Warning",
                message="Couldn't connect."
            )
            return
        if i["name"] not in os.listdir("scripts"):
            update = True
            with open(f"scripts/{i['name']}", "w", encoding="utf-8") as f:
                f.write(file)
        else:
            with open(f"scripts/{i['name']}", "r", encoding="utf-8") as f:
                if file != f.read():
                    update = True
                    with open(
                            f"scripts/{i['name']}",
                            "w",
                            encoding="utf-8"
                    ) as g:
                        g.write(file)
    if update:
        showinfo(
            title="Info",
            message="Program is updated."
        )
        if os.path.exists("defaults.ini"):
            os.remove("defaults.ini")
        if os.name == "posix":
            Popen(["python3", "WFM-BPGen.py"])
            os.kill(os.getpid(), __import__("signal").SIGKILL)
        elif os.name == "nt":
            Popen(["python", "WFM-BPGen.py"])
            os.system(f"TASKKILL /F /PID {os.getpid()}")
    else:
        showinfo(
            title="Info",
            message="Program is up-to-date."
        )


class About(tk.Toplevel):
    def __init__(self, version=__version__):
        super().__init__()
        self.resizable(width=False, height=False)
        if os.name == "nt" and os.path.exists("images/pyc.ico"):
            self.wm_iconbitmap("images/pyc.ico")
        self.version = version
        self.name = "WFM-BPGen"
        self.date_built = "01.05.2022"
        self.date_updated = "02.08.2022"
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
