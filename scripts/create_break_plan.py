# -*- coding: utf-8 -*-

from .libs import (
    os, json, np, pd, perf_counter, Workbook, harmonic_mean, dt, td
)

BREAKS = {
    "Rest-1": {
        "Start": 1,
        "End": 2,
        "Minutes": 15
    },
    "Lunch": {
        "Start": 3,
        "End": 5.25,
        "Minutes": 45
    },
    "Rest-2": {
        "Start": 6.25,
        "End": 7.75,
        "Minutes": 15
    },
    "Quiz": {
        "Start": 0.5,
        "End": 7.75,
        "Minutes": 30
    },
    "Wellness 1": {
        "Start": 3,
        "End": 7.75,
        "Minutes": 15
    },
    "Wellness 2": {
        "Start": 3,
        "End": 7.75,
        "Minutes": 15
    }
}


def counter(shift, date, data, previous=False):
    if isinstance(shift, str) and len(shift) == 11 and "-" in shift:
        start, end = shift.split("-")
        start = dt.strptime(f"{date} {start}", "%m/%d/%Y %H:%M")
        end = dt.strptime(f"{date} {end}", "%m/%d/%Y %H:%M")
        if start.hour >= 15:
            end += td(days=1)
        for i in range(24):
            hour = dt.strptime(date, "%m/%d/%Y")
            if previous:
                hour += td(days=1)
            hour += td(hours=i)
            if start <= hour < end:
                data[i] += 1


def get_hc(df, date):
    previous_date = dt.strptime(date, "%m/%d/%Y") - td(days=1)
    previous_date = previous_date.strftime("%m/%d/%Y")
    data = []
    columns = ["Skill", "Time", "HC"]
    for skill in sorted(set(df["Skill"])):
        skill_data = {i: 0 for i in range(24)}
        for shift in df[df["Skill"] == skill][date]:
            counter(shift=shift, date=date, data=skill_data)
        for shift in df[df["Skill"] == skill][previous_date]:
            counter(
                shift=shift,
                date=previous_date,
                data=skill_data,
                previous=True
            )
        data.extend(
            [
                [skill, key, values]
                for key, values in skill_data.items()
            ]
        )
    return pd.DataFrame(data=data, columns=columns)


def get_avg_values_of_n_days(filename, skill, n_days):
    if skill in ["TTR1_Recalled", "TTR2_Label1", "TTR2_Long"]:
        _input = "No.of High-pri. Incoming"
        _output = "No.of High-pri. Output"
    else:
        _input = "No.of Total Input"
        _output = "No.of Total Output"
    df = pd.read_excel(filename, sheet_name=skill)
    date = sorted(set(df["Date"]))[-1]
    start = (date - td(days=n_days)).strftime("%m/%d/%Y")
    end = date.strftime("%m/%d/%Y")
    data1 = []
    columns1 = [f"No.of Total Input (Avg.of {start} - {end}"]
    data2 = []
    columns2 = [f"AHT (s) (Avg.of {start} - {end}"]
    df = df[df["Date"] >= start]
    df = df[df["Date"] <= end]
    df["Time"] = [int(i.strftime("%H")) for i in df["Time"]]
    for i in range(24):
        noti = df[df["Time"] == i][f"{_input}"]
        aht = df[df["Time"] == i][f"AHT (s)"]
        data1.append(sum(noti) / len(noti))
        data2.append(sum(noti * aht) / sum(noti))
    return (
        pd.DataFrame(data=data1, columns=columns1),
        pd.DataFrame(data=data2, columns=columns2)
    )


def get_hourly_trend(avg_input):
    data = avg_input.values / sum(avg_input.values)
    return pd.DataFrame(data=data, columns=["Hourly Trend"])


def get_hourly_forecast(avg_input):
    data = avg_input.values
    return pd.DataFrame(data=data, columns=["Hourly Forecast"])


def get_need(avg_input, aht_target, shrinkage):
    data = (avg_input.values * aht_target.values) / (3600 * shrinkage)
    return pd.DataFrame(data=data, columns=["Need"])


def get_targets(filename, columns):
    df = pd.read_excel(filename, sheet_name="Map")[columns]
    return {
        i: {
            "AHT Target": round(
                df[df["LoB"] == i]["AHT Target"].values.tolist()[0],
                2
            )
        }
        for i in df["Skill"]
        if pd.notna(i)
    }


def get_intervals(filename, progress=None):
    columns = ["Skill", "AHT Target", "LoB"]
    targets = get_targets(filename=filename, columns=columns)
    data = pd.DataFrame()
    size = len(targets)
    now = perf_counter()
    received = 0
    for skill in targets:
        received += 1
        progress(
            s=size,
            r=received,
            n=now
        )
        avg_input, avg_aht = get_avg_values_of_n_days(
            filename=filename,
            skill=skill,
            n_days=7
        )
        need = get_need(
            avg_input=avg_input,
            aht_target=avg_aht,
            shrinkage=0.9
        )
        need = need.assign(
            **{"Skill": [skill] * 24, "Time": [*range(24)]}
        )
        data = data.append(need)
    return data


def get_shift_plan(filename):
    df = pd.read_excel(io=filename, sheet_name="Overall")
    df = pd.DataFrame(data=df.values[42:], columns=df.values[41])
    columns = ["Username", "Manager", "Skill", "Name Surname"]
    columns += [i for i in df.columns if isinstance(i, dt)]
    df = df[columns]
    df.columns = [
        i if isinstance(i, str) else i.strftime("%m/%d/%Y")
        for i in df.columns
    ]
    return df


def get_15_minutes_interval(df, date):
    data = []
    for index, row in enumerate(df.values):
        for i in range(4):
            t = dt.strptime(f"{date} {row[1]}:{i * 15}", "%m/%d/%Y %H:%M")
            if t.hour == 8 and i in [2, 3]:
                hc = df.values[index + 1][2]
            else:
                hc = df.values[index][2]
            data.append([row[0], t, hc, row[3]])
    return pd.DataFrame(data=data, columns=["Skill", "Date", "HC", "Need"])


def get_shift_times(date, shift):
    start, end = shift.split("-")
    start = dt.strptime(f"{date} {start}", "%m/%d/%Y %H:%M")
    end = dt.strptime(f"{date} {end}", "%m/%d/%Y %H:%M")
    if start.hour > 14:
        end += td(days=1)
    return start, end


def convert_skillname(shift_plan, hc):
    skill_conversions = {
        i: j
        for i, j in zip(
            sorted(set(shift_plan["Skill"])),
            sorted(set(hc["Skill"]))
        )
    }
    skill_conversions["Audio"] = "Audio_TR"
    skill_conversions["Audio-AR"] = "Audio_AR"
    return skill_conversions


def read_json(filename, data=None, breaks=True):
    if not data:
        if breaks:
            data = BREAKS
        else:
            data = {}
    if os.path.exists(filename):
        with open(filename) as f:
            return json.load(f)
    else:
        write_json(filename, data)
        return read_json(filename)


def write_json(filename, data):
    with open(filename, "w") as f:
        json.dump(data, f, indent=4, ensure_ascii=True)


def has_conflict(breaks, break_start_time, break_times, key):
    for k, v in break_times.items():
        if v:
            minutes = int(breaks[k]["Minutes"])
            for i in range(minutes // 15):
                if break_start_time == v + td(minutes=15 * i):
                    return True
            minutes = breaks[key]["Minutes"]
            for i in range(minutes // 15):
                if break_start_time == v - td(minutes=15 * i):
                    return True
    return False


def create_break_plan(intervals, date, shift_plan, breaks, progress):
    date = dt.strptime(date, "%m/%d/%Y")
    date2 = (date + td(days=1)).strftime("%m/%d/%Y")
    date1 = date.strftime("%m/%d/%Y")
    hc1 = get_15_minutes_interval(df=intervals, date=date1)
    hc2 = get_15_minutes_interval(df=intervals, date=date2)
    skill_conversions = convert_skillname(shift_plan=shift_plan, hc=hc1)
    data1 = shift_plan[["Username", "Manager", "Skill", "Name Surname", date1]]
    left1 = {}
    left2 = {}
    break_plan = []
    skills = sorted(set(data1["Skill"]))
    size = len(skills)
    received = 0
    now = perf_counter()
    quiz = read_json(filename="quiz.json", data={}, breaks=False)
    for skill in skills:
        received += 1
        progress(
            s=size,
            r=received,
            n=now
        )
        data_skill1 = data1[data1["Skill"] == skill]
        hc_skill1 = hc1[hc1["Skill"] == skill_conversions[skill]]
        hc_skill2 = hc2[hc2["Skill"] == skill_conversions[skill]]
        left1[skill] = {
            (dt.strptime(date1, "%m/%d/%Y") + td(minutes=i * 15)): 0
            for i in range(24 * 4)
        }
        left2[skill] = {
            (dt.strptime(date2, "%m/%d/%Y") + td(minutes=i * 15)): 0
            for i in range(24 * 4)
        }
        for shift in sorted(
                set(
                    [
                        i for i in data_skill1[date1]
                        if isinstance(i, str) and len(i) == 11 and i[0].isnumeric()
                    ]
                )
        ):
            data = data_skill1[data_skill1[date1] == shift]
            start, end = get_shift_times(date=date1, shift=shift)
            shuffled = data.values
            np.random.shuffle(shuffled)
            for people in shuffled:
                break_times = {}
                for key, value in breaks.items():
                    if key in break_times:
                        continue
                    if key == "Quiz":
                        if date.isoweekday() == 1 and shift in ["00:00-08:30", "08:00-17:00"]:
                            break_times["Quiz"] = ""
                            continue
                        if people[-2] not in quiz:
                            quiz[people[-2]] = True
                        else:
                            break_times["Quiz"] = ""
                            continue
                    alternatives = {}
                    alternative_minutes = int(
                        ((value["End"] - value["Start"]) * 60) / 15
                    )
                    for i in range(0, alternative_minutes):
                        break_start_time = (
                            start
                            +
                            td(hours=value["Start"])
                            +
                            td(minutes=15 * i)
                        )
                        if has_conflict(
                            breaks=breaks,
                            break_times=break_times,
                            break_start_time=break_start_time,
                            key=key
                        ):
                            continue
                        ps = []
                        for interval in range(value["Minutes"] // 15):
                            int_time = break_start_time + td(minutes=interval * 15)
                            if int_time >= dt.strptime(date2, "%m/%d/%Y"):
                                hc_skill = hc_skill2
                                left = left2
                            else:
                                hc_skill = hc_skill1
                                left = left1
                            hc_skill_date = hc_skill[hc_skill["Date"] == int_time]
                            hc_skill_total = hc_skill_date["HC"].values[0]
                            hc_skill_need = hc_skill_date["Need"].values[0]
                            p = ((hc_skill_total - left[skill][int_time]) / hc_skill_need) * 100
                            ps += [p]
                        alternatives[break_start_time] = harmonic_mean(ps)
                    p_values = [value for value in alternatives.values()]
                    index = p_values.index(max(p_values))
                    for j in range(value["Minutes"] // 15):
                        t = [*alternatives.keys()][index] + td(minutes=j * 15)
                        if j == 0:
                            break_times[key] = t
                        if t >= dt.strptime(date2, "%m/%d/%Y"):
                            left2[skill][t] += 1
                        else:
                            left1[skill][t] += 1
                break_plan.append(
                    [
                        dt.strptime(date1, "%m/%d/%Y"),
                        *people[:-1],
                        people[-1],
                        *break_times.values()
                    ]
                )
    left_data = [
        [skill, value, left1[skill][value]]
        for skill in left1 for value in left1[skill]
    ]
    left_columns = ["Skill", "Minute", "Left"]
    left_df = pd.DataFrame(data=left_data, columns=left_columns)
    hc = hc1.assign(**{"Left": left_df["Left"].values})
    hc = hc.assign(**{"Remaining": hc["HC"].values - hc["Left"].values})
    hc = hc.assign(**{"Percentage": hc["Remaining"] / hc["Need"] * 100})
    hc["Need"] = [round(i, 2) for i in hc["Need"]]
    hc["Percentage"] = [round(i, 2) for i in hc["Percentage"].values]
    columns = [
        "Date",
        "Username",
        "Manager",
        "Skill",
        "Name Surname",
        "Shift",
        *breaks.keys()
    ]
    b_plan = pd.DataFrame(data=break_plan, columns=columns)
    b_plan = b_plan[b_plan["Date"] == date]
    columns = [
        "Date",
        "Manager",
        "Skill",
        "Username",
        "Name Surname",
        "Shift",
        *breaks.keys()
    ]
    b_plan = b_plan[columns]
    write_json(filename="quiz.json", data=quiz)
    to_excel(date, b_plan, hc)


def break_plan_format(j):
    if j == 0:
        return {
            "align": "center",
            "num_format": "d.mm.YYYY"
        }
    elif j == 4:
        return {
            "bg_color": "#ffffff",
            "color": "#000000",
            "align": "left"
        }
    elif j == 5:
        return {
            "bg_color": "#66cc00",
            "color": "#ffffff",
            "align": "center",
            "border": 1
        }
    elif j >= 6:
        return {
            "align": "center",
            "num_format": "hh:mm"
        }
    else:
        return {"align": "center"}


def interval_format(j):
    if j == 1:
        return {
            "align": "center",
            "num_format": "hh:mm"
        }
    else:
        return {"align": "center"}


def to_excel(date, break_plan, hc):
    date_format = f"{date.month}.{date.day}"
    filename = f"{date_format} Break Plan.xlsx"
    cell_format = {
        "bg_color": "#4b286d",
        "color": "#ffffff",
        "align": "center",
        "bold": 1,
        "border": 1
    }
    with Workbook(filename=filename) as wb:
        for index, (name, data) in enumerate(
            [
                [f"{date_format} Break Plan", break_plan],
                [f"{date_format} Intervals", hc]
            ]
        ):
            ws = wb.add_worksheet(name=name)
            ws.hide_gridlines(2)
            for i, column in enumerate(data.columns):
                ws.write(0, i, column, wb.add_format(cell_format))
            ws.autofilter(
                0,
                0,
                len(data.values),
                len(data.columns) - 1
            )
            for i, row in enumerate(data.values):
                for j, column in enumerate(data.values[i]):
                    if index == 0:
                        frmt = break_plan_format(j=j)
                    else:
                        frmt = interval_format(j=j)
                    try:
                        ws.write(i + 1, j, column, wb.add_format(frmt))
                    except (ValueError, TypeError):
                        pass
