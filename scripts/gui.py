# -*- coding: utf-8 -*-

from .libs import (
    os, xlrd, tk, ttk, Thread, perf_counter,
    Image, ImageTk, Calendar, askopenfilename,
    showinfo, showwarning, showerror, pd, dt
)
from .about import About
from .create_break_plan import (
    get_shift_plan, get_intervals,
    get_hc, create_break_plan, read_json, write_json
)


def create_images():
    images = {}
    for img in os.listdir("./images"):
        filename = os.path.splitext(os.path.split(img)[-1])[0]
        path = os.path.join(os.getcwd(), "Images", img)
        img = Image.open(path)
        img = ImageTk.PhotoImage(img)
        images[filename] = img
    return images


class Combobox(ttk.Combobox):
    def __init__(self, option, text, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.var = tk.IntVar()
        self.var.set(option[text])
        self["textvariable"] = self.var
        self["state"] = "readonly"
        self["width"] = 10


class ComboboxToplevel(tk.Toplevel):
    def __init__(self, images, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("Edit Defaults")
        self.resizable(width=False, height=False)
        self.images = images
        self.breaks = read_json("defaults.json")
        if os.name == "nt" and os.path.exists("images/pyc.ico"):
            self.wm_iconbitmap("images/pyc.ico")
        self.values = {
            "hours": [i * 0.25 for i in range(37)],
            "minutes": [i * 15 for i in range(37)]
        }
        self.frame = tk.Frame(master=self)
        self.frame.pack()
        self.activities = {}
        self.widget = self.create_widgets(options=self.breaks)
        self.button_frame = tk.Frame(master=self)
        self.button_frame.pack(pady=10)
        self.add_button = tk.Button(
            master=self.button_frame,
            text="Add",
            compound="top",
            image=self.images["add"],
            highlightthickness=0,
            borderwidth=0,
            command=lambda: self.add_command()
        )
        self.add_button.pack(side="left", padx=10)
        self.remove_button = tk.Button(
            master=self.button_frame,
            text="Remove",
            compound="top",
            image=self.images["remove"],
            highlightthickness=0,
            borderwidth=0,
            command=self.remove_command
        )
        self.remove_button.pack(side="left", padx=10)
        self.button = tk.Button(
            master=self,
            text="OK",
            image=self.images["tick"],
            compound="left",
            highlightthickness=0,
            borderwidth=0,
            command=self.apply
        )
        self.button.pack(pady=10)

    def apply(self):
        for key, value in self.widget.items():
            start = value["Start"].get()
            if not start:
                showwarning(
                    title="Warning",
                    message=f"Select the start value for {key}."
                )
                return
            end = value["End"].get()
            if not end:
                showwarning(
                    title="Warning",
                    message=f"Select the end value for {key}."
                )
                return
            minutes = value["Minutes"].get()
            if not minutes:
                showwarning(
                    title="Warning",
                    message=f"Select the minutes value for {key}."
                )
                return
        data = {
            key: {
                subkey: (
                    int(self.widget[key][subkey].get())
                    if subkey == "Minutes"
                    else
                    float(self.widget[key][subkey].get())
                )
                for subkey in self.widget[key]
                if subkey not in ["Checkbutton", "Label", "var"]
            } for key in self.widget
        }
        write_json(filename="defaults.json", data=data)
        self.destroy()

    def create_widgets(self, options):
        widget = {}
        for index, option in enumerate(options):
            row = tk.Label(
                master=self.frame,
                text=option,
                font="Default 9 bold"
            )
            row.grid(row=index + 1, column=0, sticky="w")
            if index < 3:
                column = tk.Label(
                    master=self.frame,
                    text=list(options[option].keys())[index],
                    font="Default 9 bold"
                )
                column.grid(row=0, column=index + 1, sticky="w")
            start, end, minutes, checkbutton, var = self.triple_widgets(
                options=options,
                option=option,
                index=index
            )
            widget[option] = {
                "Start": start,
                "End": end,
                "Minutes": minutes,
                "Checkbutton": checkbutton,
                "var": var,
                "Label": row
            }
        return widget

    def triple_widgets(self, options, option, index):
        start = Combobox(
            master=self.frame,
            option=options[option],
            values=self.values["hours"],
            text="Start"

        )
        start.grid(row=index + 1, column=1, sticky="w")
        end = Combobox(
            master=self.frame,
            option=options[option],
            values=self.values["hours"],
            text="End"
        )
        end.grid(row=index + 1, column=2, sticky="w")
        minutes = Combobox(
            master=self.frame,
            option=options[option],
            values=self.values["minutes"],
            text="Minutes"
        )
        minutes.grid(row=index + 1, column=3, sticky="w")
        var = tk.BooleanVar()
        checkbutton = tk.Checkbutton(master=self.frame, variable=var)
        checkbutton.grid(row=index + 1, column=4, sticky="w")
        return start, end, minutes, checkbutton, var

    def add_command(self):
        activity_name = []
        self.ask_name(activity_name=activity_name)
        if not activity_name:
            return
        label = tk.Label(master=self.frame, text=activity_name[0], font="Default 9 bold")
        label.grid(row=len(self.frame.winfo_children()), column=0, sticky="w")
        start, end, minutes, checkbutton, var = self.triple_widgets(
            options={activity_name[0]: {"Start": "", "End": "", "Minutes": ""}},
            option=activity_name[0],
            index=len(self.frame.winfo_children()) - 1
        )
        self.widget[activity_name[0]] = {
            "Label": label,
            "Start": start,
            "End": end,
            "Minutes": minutes,
            "Checkbutton": checkbutton,
            "var": var
        }

    def remove_command(self):
        to_be_deleted = []
        for key, value in self.widget.items():
            if value["var"].get():
                for widget in value:
                    if widget != "var":
                        value[widget].destroy()
                to_be_deleted += [key]
        for activity in to_be_deleted:
            self.widget.pop(activity)

    def ask_name(self, activity_name):
        toplevel = tk.Toplevel()
        label = tk.Label(master=toplevel, text="Offline Activity Name")
        label.pack()
        entry = ttk.Entry(master=toplevel)
        entry.pack()
        button = tk.Button(
            master=toplevel,
            text="OK",
            image=self.images["tick"],
            compound="left",
            highlightthickness=0,
            borderwidth=0,
            command=lambda: self.apply_command(
                toplevel=toplevel,
                entry=entry,
                activity_name=activity_name
            )
        )
        button.pack()
        toplevel.wait_window()
        return activity_name

    def apply_command(self, toplevel, entry, activity_name):
        value = entry.get()
        if value:
            if value in self.activities or value in self.widget:
                showwarning(
                    title="Warning",
                    message=f"There's an offline activity with this name."
                )
                return
            toplevel.destroy()
            activity_name += [value]
        else:
            showwarning(
                title="Warning",
                message=f"You didn't specify a name."
            )
            return


class Menu(tk.Menu):
    def __init__(self, texts, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.active = False
        self.data = {}
        self.filename = {}
        self.breaks = read_json(filename="defaults.json")
        tk.Frame(master=self.master).pack(pady=5)
        self.label_frame = tk.Frame(master=self.master)
        self.label_frame.pack(fill="both", padx=5)
        tk.Frame(master=self.master).pack(pady=2.5)
        self.frame = tk.Frame(master=self.master)
        self.frame.pack()
        self.images = create_images()
        self.file = tk.Menu(master=self, tearoff=False)
        self.file_open = tk.Menu(master=self.file, tearoff=False)
        self.add_cascade(label="File", menu=self.file)
        self.add_command(
            label="Edit",
            command=lambda: ComboboxToplevel(
                images=self.images,
            )
        )
        self.add_command(
            label="About",
            command=lambda: About(
            )
        )
        self.file.add_cascade(
            image=self.images["open"],
            label="Open",
            compound="left",
            menu=self.file_open
        )
        tk.Frame(master=self.frame).pack(pady=5)
        self.widget = self.create_labels(texts=texts)
        for text in texts:
            self.file_open.add_command(
                label=text,
                command=lambda t=text: self.open_file(title=t)
            )
        self.calendar = Calendar(master=self.frame)
        self.calendar.pack()
        tk.Frame(master=self.frame).pack(pady=2.5)
        self.button = tk.Button(
            master=self.frame,
            text="Start",
            image=self.images["play"],
            command=lambda: Thread(target=self.start, daemon=True).start(),
            compound="left",
            highlightthickness=0,
            borderwidth=0
        )
        self.button.pack()
        tk.Frame(master=self.frame).pack(pady=2.5)
        self.progressframe = tk.Frame(master=self.frame)
        self.progresslabelframe = tk.Frame(master=self.frame)
        self.progresslabelframe.pack()
        self.vars = {}
        for i, (text, unit) in enumerate(
                [("Status", "%"), ("ETA", "s"), ("Part", "")]
        ):
            if text == "Part":
                val = "0/3"
            else:
                val = "0"
            var = tk.StringVar()
            var.set(val)
            self.vars[text] = var
            title = tk.Label(
                master=self.progresslabelframe,
                text=text,
                font="Default 9 bold"
            )
            title.grid(row=i, column=0, sticky="w")
            colon = tk.Label(
                master=self.progresslabelframe,
                text=" : ",
                font="Default 9 bold"
            )
            colon.grid(row=i, column=1)
            value = tk.Label(
                master=self.progresslabelframe,
                textvariable=var,
                font="Default 8"
            )
            value.grid(row=i, column=2, sticky="w")
            unit = tk.Label(
                master=self.progresslabelframe,
                text=unit,
                font="Default 8"
            )
            unit.grid(row=i, column=3)
        tk.Frame(master=self.progressframe).pack(pady=2.5)
        self.progressbar = ttk.Progressbar(
            master=self.progressframe,
            orient="horizontal",
            mode="determinate",
            length=200
        )
        self.progressbar.pack()
        self.progressframe.pack()

    def progress(self, s, r, n):
        self.progressbar["value"] = r
        self.progressbar["maximum"] = s
        self.vars["Status"].set(f"{int(100 * r / s)}")
        self.vars["ETA"].set(
            "{}".format(
                round(
                    (
                        int(s / (r / (perf_counter() - n)))
                        -
                        int(perf_counter() - n)
                    )
                )
            )
        )

    def start(self):
        if self.active:
            self.master.after(
                0,
                lambda: showwarning(
                    title="Warning",
                    message=f"Another process is still running.",
                    parent=self.master
                )
            )
            return
        self.active = True
        date = self.calendar.get_date()
        date = dt.strptime(date, "%m/%d/%y").strftime("%m/%d/%Y")
        shift_plan = self.widget["Shift Plan"]["text"]
        intervals = self.widget["RTA LoB Hourly Status"]["text"]
        for title, filename in [
            ("Shift Plan", shift_plan),
            ("RTA LoB Hourly Status", intervals)
        ]:
            if filename == "None":
                self.master.after(
                    0,
                    lambda: showwarning(
                        title="Warning",
                        message=f"{title} not found.",
                        parent=self.master
                    )
                )
                self.vars["Part"].set("0/3")
                self.active = False
                return
        self.vars["Part"].set("1/3")
        if "Shift Plan" not in self.data:
            try:
                self.data["Shift Plan"] = get_shift_plan(
                    filename=self.filename["Shift Plan"]
                )
            except xlrd.biffh.XLRDError:
                self.master.after(
                    0,
                    lambda: showerror(
                        title="Error",
                        message="The file has wrong format for shift plan.",
                        parent=self.master
                    )
                )
                self.vars["Part"].set("0/3")
                self.active = False
                return
        if "RTA LoB Hourly Status" not in self.data:
            try:
                if os.path.exists("intervals.xlsx"):
                    df = pd.read_excel("intervals.xlsx")
                    df = df[df.columns[1:]]
                else:
                    df = get_intervals(
                        filename=self.filename["RTA LoB Hourly Status"],
                        progress=self.progress
                    )
                    df.to_excel("intervals.xlsx")
                self.data["RTA LoB Hourly Status"] = df
            except xlrd.biffh.XLRDError:
                self.master.after(
                    0,
                    lambda: showerror(
                        title="Error",
                        message="The file has wrong format for RTA LoB Hourly Status.",
                        parent=self.master
                    )
                )
                self.vars["Part"].set("0/3")
                self.active = False
                return
        hc = get_hc(df=self.data["Shift Plan"], date=date)
        intervals = self.data["RTA LoB Hourly Status"]
        intervals = intervals.assign(**{"HC": hc["HC"].values})
        intervals = intervals[["Skill", "Time", "HC", "Need"]]
        breaks = read_json("defaults.json")
        self.vars["Part"].set("2/3")
        create_break_plan(
            intervals=intervals,
            shift_plan=self.data["Shift Plan"],
            date=date,
            breaks=breaks,
            progress=self.progress,
            var=self.vars
        )
        self.active = False
        self.master.after(
            0,
            lambda: showinfo(
                title="Info",
                message=f"The break plan has been successfully created.",
                parent=self.master
            )
        )

    def create_labels(self, texts):
        widget = {}
        for index, text in enumerate(texts):
            title = tk.Label(master=self.label_frame, text=text, font="Default 9 bold")
            title.grid(row=index, column=0, sticky="w")
            colon = tk.Label(master=self.label_frame, text=":", font="Default 9 bold")
            colon.grid(row=index, column=1, sticky="w")
            label = tk.Label(master=self.label_frame, text="None", font="Default 8")
            label.grid(row=index, column=2, sticky="w")
            widget[text] = label
        return widget

    def open_file(self, title):
        try:
            filename = askopenfilename(
                title=title,
                filetypes=[("Excel Files", "*.xlsx")]
            )
        except FileNotFoundError:
            showwarning(
                title="Warning",
                message=f"{title} Not Found.",
                parent=self.master
            )
            return
        if filename:
            self.filename[title] = filename
            self.widget[title]["text"] = os.path.split(filename)[-1]


def main():
    root = tk.Tk()
    root.title("WFM-BPGen")
    if os.name == "nt" and os.path.exists("images/pyc.ico"):
        root.wm_iconbitmap("images/pyc.ico")
    root.geometry("400x400")
    root.resizable(width=False, height=False)
    menu = Menu(master=root, texts=["Shift Plan", "RTA LoB Hourly Status"])
    root.config(menu=menu)
    root.mainloop()


if __name__ == "__main__":
    main()
