# -*- coding: utf-8 -*-

from .libs import (
    os, json, np, pd, perf_counter, Workbook, dt, td
)

BREAKS = {
    "Rest-1": {
        "Start": 1,
        "End": 3,
        "Minutes": 15
    },
    "Lunch": {
        "Start": 3,
        "End": 6,
        "Minutes": 45
    },
    "Rest-2": {
        "Start": 6,
        "End": 8,
        "Minutes": 15
    },
    "Quiz": {
        "Start": 0.5,
        "End": 8,
        "Minutes": 30
    },
    "Wellness 1": {
        "Start": 3,
        "End": 6,
        "Minutes": 15
    },
    "Wellness 2": {
        "Start": 6,
        "End": 8,
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
    if previous_date in df.columns:
        previous = True
    else:
        previous = False
    data = []
    columns = ["Skill", "Time", "HC"]
    skills = sorted(set(df["Skill"]), key=str.lower)
    for skill in skills:
        skill_data = {i: 0 for i in range(24)}
        for shift in df[df["Skill"] == skill][date]:
            counter(shift=shift, date=date, data=skill_data)
        if previous:
            for shift in df[df["Skill"] == skill][previous_date]:
                counter(
                    shift=shift,
                    date=previous_date,
                    data=skill_data,
                    previous=previous
                )
        data.extend(
            [
                [skill, key, values]
                for key, values in skill_data.items()
            ]
        )
    return pd.DataFrame(data=data, columns=columns)


def get_avg_values_of_n_days(filename, skill, n_days):
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
        try:
            data1.append(sum(noti) / len(noti))
            data2.append(sum(noti * aht) / sum(noti))
        except ZeroDivisionError:
            data1.append(0)
            data2.append(0)
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


def get_skills(filename):
    df = pd.read_excel(filename, sheet_name="Map")["Skill"]
    return sorted(
        [i for i in df.values if pd.notna(i)],
        key=str.lower
    )


def get_intervals(filename, progress=None):
    skills = get_skills(filename=filename)
    data = pd.DataFrame()
    size = len(skills)
    now = perf_counter()
    received = 0
    for skill in skills:
        received += 1
        print(skill)
        progress(
            s=size,
            r=received,
            n=now
        )
        avg_input, avg_aht = get_avg_values_of_n_days(
            filename=filename,
            skill=skill,
            n_days=2
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
    df = pd.DataFrame(data=df.values[44:], columns=df.values[43])
    columns = ["Aze User", "Manager", "Skill", "Name Surname"]
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
            data.append([row[0], t, hc, *row[3:]])
    df = pd.DataFrame(
        data=data,
        columns=["Skill", "Datetime", "HC", "Need"]
    )
    df = df.assign(**{"Left": [0] * len(df)})
    df = df.assign(**{"Remaining": df["HC"].values})
    return df


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
            sorted(set(shift_plan["Skill"]), key=str.lower),
            sorted(set(hc["Skill"]), key=str.lower)
        )
    }
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


def has_long_break_gaps(key, breaks, break_start_time, break_times, minimum, maximum):
    if key == "Rest-1":
        return False
    if key not in ["Lunch", "Rest-2"]:
        return False
    previous_key = list(breaks)[list(breaks).index(key) - 1]
    previous_break_time = break_times[previous_key]
    if break_start_time >= previous_break_time + td(hours=maximum):
        return True
    if break_start_time <= previous_break_time + td(hours=minimum):
        return True
    return False


def sort_break_plan(df):
    column_index = [i for i in df.columns]
    mandatory_activities = ["Rest-1", "Lunch", "Rest-2"]
    for skill in sorted(set(df["Skill"])):
        df_skill = df[df["Skill"] == skill]
        for shift in sorted(set(df_skill["Shift"])):
            df_skill_shift = df_skill[df_skill["Shift"] == shift]
            indexes = [*df_skill_shift.index]
            for i in mandatory_activities:
                data = df_skill_shift.sort_values(by=i)
                df.iloc[indexes[0]:indexes[-1] + 1, column_index.index(i)] = data[i].values
    return df


def create_break_plan(
        intervals,
        date,
        shift_plan,
        breaks,
        progress,
        var,
        recursion=False,
        hc1=None,
        hc2=None,
        b_plan=None
):
    tomorrow = (dt.strptime(date, "%m/%d/%Y") + td(days=1)).strftime("%m/%d/%Y")
    if not recursion:
        yesterday = dt.strptime(date, "%m/%d/%Y") - td(days=1)
        filename = f"{yesterday.month}.{yesterday.day} Break Plan.xlsx"
        breaks = {k: v for k, v in breaks.items() if k in ["Rest-1", "Lunch", "Rest-2"]}
        hc1 = get_15_minutes_interval(df=intervals, date=date)
        skill_conversions = convert_skillname(shift_plan=shift_plan, hc=hc1)
        hc2 = get_15_minutes_interval(df=intervals, date=tomorrow)
        if os.path.exists(filename):
            yesterday_plan = pd.read_excel(filename)
            shifts = ["17:30-02:00", "18:30-03:00", "19:30-04:00", "20:30-05:00"]
            yesterday_plan = yesterday_plan[yesterday_plan["Shift"].isin(shifts)]
            for skill in sorted(set(yesterday_plan["Skill"]), key=str.lower):
                y_plan_skill = yesterday_plan[yesterday_plan["Skill"] == skill]
                for people in y_plan_skill["Name Surname"]:
                    people_data = y_plan_skill[y_plan_skill["Name Surname"] == people]
                    for b in BREAKS.keys():
                        b_time = people_data[b].values[0]
                        if not pd.isna(b_time):
                            b_time = pd.to_datetime(b_time)
                            if b_time.day == (yesterday + td(days=1)).day:
                                s = skill_conversions[skill]
                                hc1_skill = hc1[hc1["Skill"] == s]
                                t_index = hc1_skill[hc1_skill["Datetime"] == b_time].index[0]
                                hc1.iloc[t_index, 5] -= 1
                                hc1.iloc[t_index, 4] += 1
        if dt.strptime(date, "%m/%d/%Y").isoweekday() == 1:
            write_json(filename="quiz.json", data={})
    quiz = read_json(filename="quiz.json", data={}, breaks=False)
    skill_conversions = convert_skillname(shift_plan=shift_plan, hc=hc1)
    data = shift_plan[["Aze User", "Manager", "Skill", "Name Surname", date]]
    break_plan = []
    skills = sorted(set(data["Skill"]), key=str.lower)
    size = len(skills)
    received = 0
    now = perf_counter()
    for skill in skills:
        received += 1
        progress(
            s=size,
            r=received,
            n=now
        )
        data_skill = data[data["Skill"] == skill]
        hc_skill1 = hc1[hc1["Skill"] == skill_conversions[skill]]
        hc_skill2 = hc2[hc2["Skill"] == skill_conversions[skill]]
        for shift in sorted(
            set(
                [
                    i for i in data_skill[date]
                    if isinstance(i, str) and len(i) == 11 and i[0].isnumeric()
                ]
            ),
        ):
            data_skill_shift = data_skill[data_skill[date] == shift]
            start, end = get_shift_times(date=date, shift=shift)
            shuffled = data_skill_shift.values
            if not recursion:
                np.random.shuffle(shuffled)
            for people in shuffled:
                if not recursion:
                    break_times = {}
                else:
                    break_times = b_plan[b_plan["Name Surname"] == people[3]]
                    break_times = {
                        "Rest-1": pd.to_datetime(break_times["Rest-1"].values[0]),
                        "Lunch": pd.to_datetime(break_times["Lunch"].values[0]),
                        "Rest-2": pd.to_datetime(break_times["Rest-2"].values[0])
                    }
                for key, value in breaks.items():
                    if key == "Quiz":
                        if (
                            dt.strptime(date, "%m/%d/%Y").isoweekday() == 1
                            and
                            shift in ["00:00-08:30", "08:00-17:00"]
                        ):
                            break_times["Quiz"] = ""
                            continue
                        if people[3] not in quiz:
                            quiz[people[3]] = True
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
                        if (
                            recursion
                            and
                            has_conflict(
                                breaks=read_json("defaults.json"),
                                break_start_time=break_start_time,
                                break_times=break_times,
                                key=key
                            )
                        ):
                            continue
                        if has_long_break_gaps(
                            breaks=read_json("defaults.json"),
                            break_start_time=break_start_time,
                            break_times=break_times,
                            key=key,
                            maximum=3,
                            minimum=2
                        ):
                            continue
                        ps = []
                        for interval in range(value["Minutes"] // 15):
                            int_time = break_start_time + td(minutes=interval * 15)
                            try:
                                hc_skill_date = hc_skill1[hc_skill1["Datetime"] == int_time]
                                hc_skill_remaining = hc_skill_date["Remaining"].values[0]
                                hc_skill_need = hc_skill_date["Need"].values[0]
                            except IndexError:
                                hc_skill_date = hc_skill2[hc_skill2["Datetime"] == int_time]
                                hc_skill_remaining = hc_skill_date["Remaining"].values[0]
                                hc_skill_need = hc_skill_date["Need"].values[0]
                            p = (hc_skill_remaining / hc_skill_need) * 100
                            ps += [p]
                        total = [1 / i if i != 0 else None for i in ps]
                        if None in total:
                            continue
                        alternatives[break_start_time] = 1 / (sum(ps) ** sum(total))
                    p_values = [value for value in alternatives.values()]
                    index = p_values.index(max(p_values))
                    for j in range(value["Minutes"] // 15):
                        t = [*alternatives.keys()][index] + td(minutes=j * 15)
                        if j == 0:
                            break_times[key] = t
                        try:
                            time_index = hc_skill1[hc_skill1["Datetime"] == t].index[0]
                            hc1.iloc[time_index, 5] -= 1
                            hc1.iloc[time_index, 4] += 1
                        except IndexError:
                            time_index = hc_skill2[hc_skill2["Datetime"] == t].index[0]
                            hc2.iloc[time_index, 5] -= 1
                            hc2.iloc[time_index, 4] += 1
                hc_skill1 = hc1[hc1["Skill"] == skill_conversions[skill]]
                hc_skill2 = hc2[hc2["Skill"] == skill_conversions[skill]]
                break_plan.append(
                    [
                        dt.strptime(date, "%m/%d/%Y"),
                        people[1],
                        people[2],
                        people[0],
                        people[3],
                        people[4],
                        *break_times.values()
                    ]
                )
    if not recursion:
        columns = [
            "Date",
            "Manager",
            "Skill",
            "Aze User",
            "Name Surname",
            "Shift",
            *breaks.keys()
        ]
        pd.DataFrame(data=break_plan, columns=columns)
        b_plan = pd.DataFrame(data=break_plan, columns=columns)
        b_plan = sort_break_plan(b_plan)
        breaks = {k: v for k, v in BREAKS.items() if k not in ["Rest-1", "Lunch", "Rest-2"]}
        var["Part"].set("3/3")
        create_break_plan(
            intervals=intervals,
            shift_plan=shift_plan,
            breaks=breaks,
            progress=progress,
            recursion=True,
            hc1=hc1,
            hc2=hc2,
            b_plan=b_plan,
            date=date,
            var=var
        )
    else:
        columns = [
            "Date",
            "Manager",
            "Skill",
            "Aze User",
            "Name Surname",
            "Shift",
            *BREAKS.keys()
        ]
        pd.DataFrame(data=break_plan, columns=columns)
        b_plan = pd.DataFrame(data=break_plan, columns=columns)
        hc1 = hc1.assign(**{"Percentage": hc1["Remaining"] / hc1["Need"] * 100})
        hc1["Percentage"] = [round(i, 2) for i in hc1["Percentage"]]
        write_json(filename="quiz.json", data=quiz)
        to_excel(date, b_plan, hc1)


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
    date = dt.strptime(date, "%m/%d/%Y")
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
