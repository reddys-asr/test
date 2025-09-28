# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
from datetime import datetime, timedelta, time
import math
import os
from collections import defaultdict

# ---------- Config / paths ----------
INPUT_XLSX = r"C:\Users\rreddytv\Desktop\NPT\Consolidated.xlsx"
OUTPUT_XLSX = r"C:\Users\rreddytv\Desktop\NPT\NPT_OutputNew.xlsx"

# Helper: round time to nearest 30-min interval start (floor)
def floor_to_30(dt):
    return dt.replace(minute=(0 if dt.minute < 30 else 30), second=0, microsecond=0)

def add_minutes(t: datetime, mins: int):
    return t + timedelta(minutes=mins)

def to_datetime_on_date(date_val, time_val):
    """
    Accepts date (datetime/date/string) and time (string/float/datetime.time);
    returns a datetime.
    If time_val is an Excel float (0-1) or pandas Timedelta/time, handle accordingly.
    """
    # normalize date
    if pd.isna(date_val):
        base_date = datetime.today().date()
    elif isinstance(date_val, (pd.Timestamp, datetime)):
        base_date = date_val.date()
    else:
        try:
            base_date = pd.to_datetime(date_val).date()
        except:
            base_date = datetime.today().date()

    # normalize time
    if pd.isna(time_val):
        t = time(0, 0)
    elif isinstance(time_val, time):
        t = time_val
    elif isinstance(time_val, (pd.Timestamp, datetime)):
        t = time_val.time()
    else:
        # try parse numeric (Excel time fraction)
        try:
            flo = float(time_val)
            # treat as fraction of day if between 0 and 1
            if 0 <= flo < 1:
                total_seconds = int(round(flo * 24 * 3600))
                hh = total_seconds // 3600
                mm = (total_seconds % 3600) // 60
                ss = total_seconds % 60
                t = time(hh, mm, ss)
            else:
                # try parse string
                t = pd.to_datetime(time_val).time()
        except Exception:
            try:
                t = pd.to_datetime(time_val).time()
            except Exception:
                t = time(0, 0)

    return datetime.combine(base_date, t)

# ---------- Read input workbook ----------
print("Loading workbook:", INPUT_XLSX)
xls = pd.read_excel(INPUT_XLSX, sheet_name=None)

# Expecting sheets named exactly (case-insensitive handling)
def get_sheet(df_dict, name):
    for k, v in df_dict.items():
        if k.strip().lower() == name.strip().lower():
            return v.copy()
    raise KeyError(f"Sheet '{name}' not found in workbook. Found: {list(df_dict.keys())}")

constraint_df = get_sheet(xls, "Constraint")
assoc_df = get_sheet(xls, "Associate_Roster")
manager_df = get_sheet(xls, "Manager_Roster")
heatmap_df = get_sheet(xls, "Schedule_heatmap")

# Normalize column names (strip)
constraint_df.columns = [c.strip() for c in constraint_df.columns]
assoc_df.columns = [c.strip() for c in assoc_df.columns]
manager_df.columns = [c.strip() for c in manager_df.columns]
heatmap_df.columns = [c.strip() for c in heatmap_df.columns]

# ---------- Preprocessing ----------
# Ensure date column in associate roster is datetime
if 'Date' in assoc_df.columns:
    assoc_df['Date'] = pd.to_datetime(assoc_df['Date']).dt.date
else:
    assoc_df['Date'] = pd.to_datetime(datetime.today().date()).date()

# Create a dictionary of meeting definitions from Constraint sheet
# Expected columns: Meeting_Name, Frequency(Daily/Weekly/Monthly), Meeting_Type, Manager_Availability, Duration, NPT_Threshold
constraint_df = constraint_df.rename(columns={c: c.strip() for c in constraint_df.columns})
meetings = []
for _, row in constraint_df.iterrows():
    m = {
        "Meeting_Name": str(row.get("Meeting_Name")).strip(),
        "Frequency": str(row.get("Frequency(Daily/Weekly/Monthly)") or row.get("Frequency") or "").strip(),
        "Meeting_Type": str(row.get("Meeting_Type") or "").strip(),
        "Manager_Availability": str(row.get("Manager_Availability") or "").strip().lower(),  # 'yes'/'no'
        "Duration": int(row.get("Duration") or 30),
        "NPT_Threshold": float(row.get("NPT_Threshold") or 0.0)
    }
    meetings.append(m)

meeting_names = [m['Meeting_Name'] for m in meetings]

# Create a meeting columns in Associate_Roster for each Meeting_Name (P..S like)
for mn in meeting_names:
    if mn not in assoc_df.columns:
        assoc_df[mn] = pd.NA

# Build a quick manager lookup from Manager_Roster (assume Manager_Roster has AA_Name & Manager col similar to Associate_Roster)
manager_list = set(manager_df['AA_Name'].dropna().astype(str).values) if 'AA_Name' in manager_df.columns else set()
# Also build mapping of associate -> manager
aa_to_manager = {}
if 'AA_Name' in assoc_df.columns and 'Manager' in assoc_df.columns:
    for _, r in assoc_df.iterrows():
        aa_to_manager[str(r['AA_Name']).strip()] = str(r.get('Manager', '')).strip()

# For fast manager availability checks, create manager roster keyed by manager name (rows may have dates)
manager_avail = {}
if 'AA_Name' in manager_df.columns:
    for _, r in manager_df.iterrows():
        name = str(r['AA_Name']).strip()
        manager_avail.setdefault(name, []).append(r)  # store rows for availability checks

# Preprocess heatmap: compute initial Revised_Staffing_buffer = Scheduled - Requirement
if 'Requirement' not in heatmap_df.columns or 'Scheduled' not in heatmap_df.columns or 'Interval' not in heatmap_df.columns:
    raise KeyError("Schedule_heatmap sheet must contain 'Requirement', 'Scheduled', and 'Interval' columns")

heatmap_df['NPT_Count'] = 0.0  # will accumulate (meeting_durations / 30)
heatmap_df['Revised_Staffing_heatmap'] = (heatmap_df['Scheduled'] - heatmap_df['NPT_Count']) - heatmap_df['Requirement']
# For quick lookup per (Date, Interval, Skill) we create a dict
# Normalize Date to date
if 'Date' in heatmap_df.columns:
    heatmap_df['Date'] = pd.to_datetime(heatmap_df['Date']).dt.date
else:
    heatmap_df['Date'] = pd.to_datetime(datetime.today().date()).date()

# Build a mapping from (Date, Interval string, Skill) -> heatmap row index
heatmap_index = {}
for idx, r in heatmap_df.iterrows():
    key = (r['Date'], str(r['Interval']).strip(), r.get('Skill'))
    heatmap_index[key] = idx

# Helper to update heatmap after scheduling meeting in an interval
def update_heatmap_for_interval(date_obj, interval_str, skill, added_meeting_minutes):
    """
    Adds meeting duration to NPT_Count (duration/30) and recalculates Revised_Staffing_heatmap.
    Returns True if after update Revised_Staffing_heatmap >= NPT_Threshold (we'll check separately), else False.
    """
    key = (date_obj, interval_str, skill)
    if key not in heatmap_index:
        return False, "No heatmap row for interval"
    idx = heatmap_index[key]
    # increment NPT_Count (duration/30)
    heatmap_df.at[idx, 'NPT_Count'] += (added_meeting_minutes / 30.0)
    heatmap_df.at[idx, 'Revised_Staffing_heatmap'] = (heatmap_df.at[idx, 'Scheduled'] - heatmap_df.at[idx, 'NPT_Count']) - heatmap_df.at[idx, 'Requirement']
    return True, None

# Helper to rollback the add (in case it violates threshold)
def rollback_heatmap_for_interval(date_obj, interval_str, skill, removed_meeting_minutes):
    key = (date_obj, interval_str, skill)
    if key not in heatmap_index:
        return False
    idx = heatmap_index[key]
    heatmap_df.at[idx, 'NPT_Count'] -= (removed_meeting_minutes / 30.0)
    heatmap_df.at[idx, 'Revised_Staffing_heatmap'] = (heatmap_df.at[idx, 'Scheduled'] - heatmap_df.at[idx, 'NPT_Count']) - heatmap_df.at[idx, 'Requirement']
    return True

# ---------- Build candidate intervals for each associate ----------
# We'll create a dict: associate row index -> list of candidate start datetimes (30-min steps) where a meeting of specific duration could start.
assoc_candidates = {}  # key: assoc index, value: list of datetime start times
for idx, row in assoc_df.iterrows():
    working_flag = int(row.get('Working', 1))
    if working_flag != 1:
        assoc_candidates[idx] = []
        continue

    # Use Date + Shift_start / Shift_end to create interval
    base_date = row['Date']
    try:
        shift_start_dt = to_datetime_on_date(base_date, row.get('Shift_start'))
        shift_end_dt = to_datetime_on_date(base_date, row.get('Shift_end'))
    except Exception:
        # fallback to whole day
        shift_start_dt = datetime.combine(base_date, time(9, 0))
        shift_end_dt = datetime.combine(base_date, time(17, 0))

    # build break intervals to exclude
    exclude_periods = []
    for label in ['lunch1_start', 'lunch1_end', 'break1_start', 'break1_end', 'break2_start', 'break2_end']:
        s = row.get(label.replace('_end', '_start')) if '_end' in label else row.get(label)
        # The row may have both start and end columns; handle pairs below
        pass

    # create explicit exclude pairs:
    pairs = [
        ('lunch1_start', 'lunch1_end'),
        ('break1_start', 'break1_end'),
        ('break2_start', 'break2_end')
    ]
    for (scol, ecol) in pairs:
        s_val = row.get(scol)
        e_val = row.get(ecol)
        if pd.notna(s_val) and pd.notna(e_val):
            sdt = to_datetime_on_date(base_date, s_val)
            edt = to_datetime_on_date(base_date, e_val)
            # ensure valid ordering
            if edt < sdt:
                edt = sdt + timedelta(minutes=30)
            exclude_periods.append((sdt, edt))

    # Build candidate starts at every 30-min mark between shift_start and shift_end (we will check meeting durations later).
    starts = []
    t = floor_to_30(shift_start_dt)
    # allow meetings to start exactly at shift_start
    while t + timedelta(minutes=30) <= shift_end_dt:  # at least 30-min slot
        # check not inside exclude periods
        inside_exclude = False
        for (es, ee) in exclude_periods:
            # if start time t is within exclude period OR meeting would overlap exclude for a 30-min minimal meeting,
            # mark excluded. We'll check meeting-specific duration later.
            if (t >= es and t < ee) or (t + timedelta(minutes=30) > es and t < ee):
                inside_exclude = True
                break
        if not inside_exclude:
            starts.append(t)
        t = t + timedelta(minutes=30)
    assoc_candidates[idx] = starts

# ---------- Scheduling engine ----------
# We'll maintain structures:
# assoc_schedules: dict assoc_idx -> dict meeting_name -> scheduled start datetime
assoc_schedules = {idx: {} for idx in assoc_df.index}
# manager_schedules: dict manager_name -> list of (date, start, duration, meeting_name, associate)
manager_schedules = defaultdict(list)
# unscheduled list
unscheduled = []

# Helper: check manager availability at a given datetime for a given duration
def is_manager_available(manager_name, start_dt, duration_mins, date_obj):
    """
    Check manager_avail using manager_df. We ensure manager has Working=1 on that date,
    and start_dt not inside breaks/lunch, and start_dt within shift_start/end.
    This is a best-effort: manager_df may have multiple rows - we'll check any row for that manager matching the date.
    """
    rows = manager_df[manager_df['AA_Name'].astype(str).str.strip() == str(manager_name).strip()]
    if rows.empty:
        return False
    for _, r in rows.iterrows():
        # match date if possible
        row_date = r.get('Date')
        if pd.notna(row_date):
            rd = pd.to_datetime(row_date).date()
            if rd != date_obj:
                continue
        # check working flag
        if int(r.get('Working', 1)) != 1:
            continue
        try:
            m_shift_start = to_datetime_on_date(date_obj, r.get('Shift_start'))
            m_shift_end = to_datetime_on_date(date_obj, r.get('Shift_end'))
        except Exception:
            continue
        if not (start_dt >= m_shift_start and (start_dt + timedelta(minutes=duration_mins)) <= m_shift_end):
            continue
        # check breaks
        clash = False
        pairs = [('lunch1_start', 'lunch1_end'), ('break1_start', 'break1_end'), ('break2_start', 'break2_end')]
        for sc, ec in pairs:
            s_val = r.get(sc); e_val = r.get(ec)
            if pd.notna(s_val) and pd.notna(e_val):
                sdt = to_datetime_on_date(date_obj, s_val)
                edt = to_datetime_on_date(date_obj, e_val)
                if not (start_dt + timedelta(minutes=duration_mins) <= sdt or start_dt >= edt):
                    clash = True
                    break
        if clash:
            continue
        return True
    return False

# Precompute groups by Shift_start for Team_Huddle
shift_groups = defaultdict(list)
for idx, row in assoc_df.iterrows():
    if int(row.get('Working', 1)) != 1:
        continue
    # use stringified shift_start for grouping
    ss = row.get('Shift_start')
    ss_dt = None
    try:
        ss_dt = to_datetime_on_date(row['Date'], ss)
        ss_key = ss_dt.time().strftime("%H:%M")
    except:
        ss_key = str(ss)
    shift_groups[ss_key].append(idx)

# First: Schedule Team_Huddle meetings (group meeting) with the distribution rules
team_huddle_name = "Team_Huddle"
# See if Team_Huddle exists in meetings definitions; if not skip
team_meet_def = next((m for m in meetings if m['Meeting_Name'].lower() == team_huddle_name.lower()), None)
if team_meet_def:
    print("Scheduling Team_Huddle meetings...")
    for ss_key, assoc_indices in shift_groups.items():
        # filter only associates (not managers) -> per your rule, Team_Huddle not for managers in Manager_Roster
        # We'll consider associates not in manager_list
        assoc_indices = [i for i in assoc_indices if assoc_df.at[i, 'AA_Name'] not in manager_list]
        if not assoc_indices:
            continue
        # For each group, we need to schedule within first 1 hour of Shift_start.
        # Determine group's shift_start datetime from first assoc
        first_idx = assoc_indices[0]
        base_date = assoc_df.at[first_idx, 'Date']
        shift_start_dt = to_datetime_on_date(base_date, assoc_df.at[first_idx, 'Shift_start'])
        # intervals: first 30-min = [shift_start, shift_start+30), second 30-min = [shift_start+30, shift_start+60)
        interval1_start = shift_start_dt
        interval2_start = shift_start_dt + timedelta(minutes=30)
        # Compute counts: In any 30-min interval schedule between 50% and 60% of them; remaining go to next interval (still within first hour).
        n = len(assoc_indices)
        # choose count in first interval as 55% rounded (best-effort within 50-60)
        first_count = int(round(n * 0.55))
        # ensure at least floor(0.5*n) and at most ceil(0.6*n)
        min_allow = math.floor(n * 0.5)
        max_allow = math.ceil(n * 0.6)
        first_count = max(min_allow, min(max_allow, first_count))
        second_count = n - first_count
        # assign random or ordered - we'll use the order as-is
        first_group = assoc_indices[:first_count]
        second_group = assoc_indices[first_count:]
        # For each in group, attempt to schedule at earliest candidate slot inside that 30-min interval
        def schedule_in_interval(target_indices, interval_start, interval_minutes=30, meeting_duration=team_meet_def['Duration'], meeting_name=team_meet_def['Meeting_Name']):
            for aidx in target_indices:
                date_obj = assoc_df.at[aidx, 'Date']
                skill = assoc_df.at[aidx, 'Skill'] if 'Skill' in assoc_df.columns else None
                # Build interval string to match heatmap interval formatting. We assume heatmap 'Interval' uses HH:MM-HH:MM or similar. We'll derive interval_str as 'HH:MM'
                # To find matching heatmap row, we'll convert the interval start to string 'HH:MM'
                # Search among candidate starts for a start within [interval_start, interval_start+interval_minutes)
                placed = False
                for cand in assoc_candidates.get(aidx, []):
                    if cand >= interval_start and cand < (interval_start + timedelta(minutes=interval_minutes)):
                        # check room for duration (cand + duration <= shift_end) - candidate generation already ensured 30-min minimal, but meeting may be >30
                        # verify no clash with other meetings already scheduled for this associate
                        clash = False
                        for _, mdict in assoc_schedules[aidx].items():
                            existing_start = mdict
                            existing_dur = 30  # no per-meeting duration stored there; but we can look up from constraint
                            # find meeting duration by its name if necessary (not needed here - we check only overlap)
                            # For safety assume existing_dur = 30 if unknown
                            existing_end = existing_start + timedelta(minutes=existing_dur)
                            new_end = cand + timedelta(minutes=meeting_duration)
                            if not (new_end <= existing_start or cand >= existing_end):
                                clash = True
                                break
                        if clash:
                            continue
                        # check NPT threshold for interval: create an interval string matching heatmap
                        # Many heatmaps use "HH:MM-HH:MM" format; we'll try building that based on cand 30-min step
                        interval_str = f"{cand.time().strftime('%H:%M')}-{(cand + timedelta(minutes=30)).time().strftime('%H:%M')}"
                        # check heatmap key presence first; if absent, try only using start time HH:MM
                        possible_keys = [(date_obj, interval_str, skill), (date_obj, cand.time().strftime('%H:%M'), skill)]
                        ok_heatmap_idx = None
                        for k in possible_keys:
                            if k in heatmap_index:
                                ok_heatmap_idx = k
                                break
                        if ok_heatmap_idx is None:
                            # fallback: try any heatmap row with same Date and Skill and Interval that starts at this time
                            # As last resort attempt to use any row with same Date and Skill (not ideal)
                            found = False
                            for k in heatmap_index.keys():
                                if k[0] == date_obj and k[2] == skill:
                                    ok_heatmap_idx = k
                                    found = True
                                    break
                            if not found:
                                # cannot validate NPT on this interval, skip
                                continue
                        # check initial Revised_Staffing_buffer > NPT_Threshold from constraint (for this meeting)
                        threshold = team_meet_def['NPT_Threshold']
                        idx_heat = heatmap_index[ok_heatmap_idx]
                        initial_revised_buffer = (heatmap_df.at[idx_heat, 'Scheduled'] - heatmap_df.at[idx_heat, 'Requirement'])
                        if initial_revised_buffer <= threshold:
                            # cannot schedule here
                            continue
                        # simulate adding meeting minutes to heatmap and verify resulting Revised_Staffing_heatmap >= threshold
                        update_heatmap_for_interval(ok_heatmap_idx[0], ok_heatmap_idx[1], ok_heatmap_idx[2], meeting_duration)
                        # check after-update value
                        new_val = heatmap_df.at[idx_heat, 'Revised_Staffing_heatmap']
                        if new_val < threshold:
                            # rollback and continue
                            rollback_heatmap_for_interval(ok_heatmap_idx[0], ok_heatmap_idx[1], ok_heatmap_idx[2], meeting_duration)
                            continue
                        # otherwise commit scheduling
                        assoc_schedules[aidx][meeting_name] = cand
                        # assign manager: Team_Huddle uses any available manager (not necessary to assign a particular manager in roster),
                        # but per spec we replicate this meeting to Managers in Manager_Roster - add to first available manager
                        # find any manager available at that time
                        assigned_manager = None
                        for mname in manager_list:
                            if is_manager_available(mname, cand, meeting_duration, date_obj):
                                assigned_manager = mname
                                manager_schedules[mname].append({
                                    'date': date_obj, 'start': cand, 'duration': meeting_duration, 'meeting_name': meeting_name, 'associate': assoc_df.at[aidx,'AA_Name']
                                })
                                break
                        placed = True
                        break
                if not placed:
                    unscheduled.append({
                        'AA_Name': assoc_df.at[aidx, 'AA_Name'],
                        'Date': date_obj,
                        'Meeting_Name': meeting_name,
                        'Reason': 'No available candidate slot in Team_Huddle interval respecting NPT/Breaks'
                    })
        # schedule first group
        schedule_in_interval(first_group, interval1_start, interval_minutes=30, meeting_duration=team_meet_def['Duration'])
        # schedule second group
        schedule_in_interval(second_group, interval2_start, interval_minutes=30, meeting_duration=team_meet_def['Duration'])

else:
    print("No Team_Huddle meeting defined in Constraint sheet or named differently. Skipping Team_Huddle scheduling.")

# Next: schedule other meetings (One-2-One and others)
print("Scheduling other meetings (One-2-One and others)...")
# We'll iterate over each associate row and each meeting (except Team_Huddle)
for idx, row in assoc_df.iterrows():
    if int(row.get('Working', 1)) != 1:
        continue
    aa_name = row.get('AA_Name')
    date_obj = row.get('Date')
    skill = row.get('Skill') if 'Skill' in assoc_df.columns else None

    # Candidate starts for this associate
    candidates = assoc_candidates.get(idx, []).copy()
    # For fairness, try earlier slots first
    candidates.sort()

    for mdef in meetings:
        mn = mdef['Meeting_Name']
        if mn.lower() == team_huddle_name.lower():
            continue  # already handled
        freq = mdef['Frequency'].strip().lower()
        # Decide whether to attempt to schedule this meeting for this associate on this row's date
        do_attempt = False
        if freq == 'daily' or freq == 'daily ':
            do_attempt = True
        elif freq == 'weekly' or freq == 'weekly ':
            # attempt weekly - we will ensure not more than 25% of same-shift associates are scheduled on this Date
            # compute current scheduled weekly for associates sharing the same shift on this date
            shift_start = row.get('Shift_start')
            # count shift peers
            shift_key = None
            try:
                ss_dt = to_datetime_on_date(date_obj, shift_start)
                shift_key = ss_dt.time().strftime('%H:%M')
            except:
                shift_key = str(shift_start)
            # total associates working with this shift_key on this date
            peers = [i for i, r in assoc_df.iterrows() if r['Date'] == date_obj and int(r.get('Working',1))==1 and \
                     (to_datetime_on_date(date_obj, r.get('Shift_start')).time().strftime('%H:%M') if pd.notna(r.get('Shift_start')) else '') == shift_key]
            total_peers = len(peers)
            # count how many of these peers already have this meeting scheduled on this date
            already = sum(1 for i in peers if mn in assoc_schedules.get(i, {}))
            # if scheduling this would make daily count > 25% of peers, skip
            if total_peers == 0:
                do_attempt = True
            else:
                if (already + 1) <= math.floor(total_peers * 0.25):
                    do_attempt = True
                else:
                    do_attempt = False
        elif freq == 'monthly' or freq == 'monthly ':
            # attempt monthly once per month; we assume this row's Date is the day to schedule in the month
            do_attempt = True
        else:
            # unknown frequency default: try
            do_attempt = True

        if not do_attempt:
            continue

        # Attempt to place meeting in any candidate slot for this associate
        placed = False
        for cand in candidates:
            # ensure meeting fits inside shift
            meeting_duration = mdef['Duration']
            if (cand + timedelta(minutes=meeting_duration)) > to_datetime_on_date(date_obj, row.get('Shift_end')):
                continue
            # ensure not overlapping lunch/breaks - candidate generation already filters 30-min overlaps; check for longer durations
            pairs = [('lunch1_start', 'lunch1_end'), ('break1_start', 'break1_end'), ('break2_start', 'break2_end')]
            clash = False
            for sc, ec in pairs:
                s_val = row.get(sc); e_val = row.get(ec)
                if pd.notna(s_val) and pd.notna(e_val):
                    sdt = to_datetime_on_date(date_obj, s_val)
                    edt = to_datetime_on_date(date_obj, e_val)
                    if not ((cand + timedelta(minutes=meeting_duration)) <= sdt or cand >= edt):
                        clash = True
                        break
            if clash:
                continue
            # ensure no clash with other meetings already scheduled for same associate
            conflict = False
            for existing_mn, existing_start in assoc_schedules[idx].items():
                # lookup duration for existing_mn
                existing_def = next((mm for mm in meetings if mm['Meeting_Name'] == existing_mn), None)
                existing_dur = existing_def['Duration'] if existing_def else 30
                existing_end = existing_start + timedelta(minutes=existing_dur)
                new_end = cand + timedelta(minutes=meeting_duration)
                if not (new_end <= existing_start or cand >= existing_end):
                    conflict = True
                    break
            if conflict:
                continue

            # Manager assignment rules
            assigned_manager_name = None
            if mdef['Meeting_Type'].strip().lower() == 'group meeting':
                # group meeting: any manager can take it if Manager_Availability is No; if Yes, must be direct manager (but group with direct manager seems odd)
                if mdef['Manager_Availability'] == 'yes':
                    assigned_manager_name = aa_to_manager.get(str(aa_name).strip(), None)
                    if assigned_manager_name is None:
                        # fall back to any available manager
                        for m in manager_list:
                            if is_manager_available(m, cand, meeting_duration, date_obj):
                                assigned_manager_name = m
                                break
                else:
                    # any available manager
                    for m in manager_list:
                        if is_manager_available(m, cand, meeting_duration, date_obj):
                            assigned_manager_name = m
                            break
            else:
                # One-2-One or individual meeting: if Manager_Availability == yes -> direct manager must attend (mirror)
                if mdef['Manager_Availability'] == 'yes':
                    direct_mgr = aa_to_manager.get(str(aa_name).strip(), None)
                    if direct_mgr and is_manager_available(direct_mgr, cand, meeting_duration, date_obj):
                        assigned_manager_name = direct_mgr
                    else:
                        # cannot place here if manager required but not available
                        # try next candidate
                        continue
                else:
                    # any manager is ok
                    for m in manager_list:
                        if is_manager_available(m, cand, meeting_duration, date_obj):
                            assigned_manager_name = m
                            break
                    if assigned_manager_name is None:
                        # if no manager available, it's still allowed? The rules say if No then any available manager can take this meeting.
                        # If none found, skip
                        continue

            # NPT / heatmap check for the interval: find heatmap interval corresponding to this cand
            interval_str = f"{cand.time().strftime('%H:%M')}-{(cand + timedelta(minutes=30)).time().strftime('%H:%M')}"
            # try lookups
            keys_try = [(date_obj, interval_str, skill), (date_obj, cand.time().strftime('%H:%M'), skill)]
            ok_k = None
            for k in keys_try:
                if k in heatmap_index:
                    ok_k = k
                    break
            if ok_k is None:
                # try any row with date & skill
                found = None
                for k in heatmap_index:
                    if k[0] == date_obj and k[2] == skill:
                        found = k
                        break
                if found:
                    ok_k = found
                else:
                    # no heatmap to validate NPT - skip this cand
                    continue
            heat_idx = heatmap_index[ok_k]
            threshold = mdef['NPT_Threshold']
            initial_revised_buffer = heatmap_df.at[heat_idx, 'Scheduled'] - heatmap_df.at[heat_idx, 'Requirement']
            if initial_revised_buffer <= threshold:
                # cannot schedule here
                continue
            # simulate add
            update_heatmap_for_interval(ok_k[0], ok_k[1], ok_k[2], meeting_duration)
            new_val = heatmap_df.at[heat_idx, 'Revised_Staffing_heatmap']
            if new_val < threshold:
                rollback_heatmap_for_interval(ok_k[0], ok_k[1], ok_k[2], meeting_duration)
                continue

            # commit schedule
            assoc_schedules[idx][mn] = cand
            # add to manager schedule mirror if manager assigned
            if assigned_manager_name:
                manager_schedules[assigned_manager_name].append({'date': date_obj, 'start': cand, 'duration': meeting_duration, 'meeting_name': mn, 'associate': aa_name})
                # Also mirror schedule into Manager_Roster by adding a row (replicating the meeting) - per spec "Also add these List of meetings to Managers in Manager_Roster"
                # We'll append to manager_df later in summary stage rather than inserting rows mid-loop.
            placed = True
            break

        if not placed:
            unscheduled.append({
                'AA_Name': aa_name,
                'Date': date_obj,
                'Meeting_Name': mn,
                'Reason': 'No valid slot respecting shift/breaks/NPT/manager availability'
            })

# ---------- Write schedules back to Associate_Roster dataframe ----------
for aidx, meetings_map in assoc_schedules.items():
    for mn, start_dt in meetings_map.items():
        # store as time string in the cell (HH:MM)
        assoc_df.at[aidx, mn] = start_dt.time().strftime("%H:%M")

# Also replicate manager meeting rows in Manager_Roster as new rows (append)
manager_added_rows = []
for mgr, lists in manager_schedules.items():
    for item in lists:
        # create a new replica row based on manager_df structure
        new_row = {}
        # try to find one existing row for this manager to copy general fields
        sample_rows = manager_df[manager_df['AA_Name'].astype(str).str.strip() == str(mgr).strip()]
        if not sample_rows.empty:
            sample = sample_rows.iloc[0].to_dict()
            # copy relevant fields
            new_row.update(sample)
            # set Date, Shift_start, Shift_end accordingly (we keep original shift values)
            new_row['Date'] = item['date']
            # We don't shift manager shift times; we just record the meeting in a separate field (Meeting_Name -> start time)
        else:
            # minimal default
            new_row = {'AA_Name': mgr, 'Date': item['date'], 'Working': 1}
        # Add meeting name column or reuse one if it exists
        # We'll add a Meeting_Scheduled column with "MeetingName@HH:MM"
        new_row['Meeting_Scheduled'] = f"{item['meeting_name']}@{item['start'].time().strftime('%H:%M')}"
        manager_added_rows.append(new_row)

if manager_added_rows:
    manager_df = pd.concat([manager_df, pd.DataFrame(manager_added_rows)], ignore_index=True, sort=False)

# ---------- Update Schedule_heatmap with NPT_Count and Revised_Staffing_heatmap already maintained ----------
# We've been updating heatmap_df in place.

# ---------- Summary report ----------
# Total meetings scheduled for each AA-Name and each Date
summary_rows = []
for idx, row in assoc_df.iterrows():
    aa_name = row.get('AA_Name')
    date_obj = row.get('Date')
    scheduled_meetings = assoc_schedules.get(idx, {})
    summary_rows.append({
        'AA_Name': aa_name,
        'Date': date_obj,
        'Total_Meetings_Scheduled': len(scheduled_meetings),
        'Meetings_List': ", ".join([f"{k}@{v.time().strftime('%H:%M')}" for k, v in scheduled_meetings.items()])
    })
summary_df = pd.DataFrame(summary_rows)

# Team Huddle distribution percentages: per shift_start group, compute actual distribution for Team_Huddle
team_huddle_stats = []
if team_meet_def:
    for ss_key, assoc_indices in shift_groups.items():
        assoc_indices = [i for i in assoc_indices if assoc_df.at[i,'AA_Name'] not in manager_list]
        total = len(assoc_indices)
        if total == 0:
            continue
        first_interval_count = 0
        # first interval start
        if total > 0:
            first_idx = assoc_indices[0]
            base = assoc_df.at[first_idx, 'Date']
            ss_dt = to_datetime_on_date(base, assoc_df.at[first_idx,'Shift_start'])
            int1_start = ss_dt
            int1_end = ss_dt + timedelta(minutes=30)
            for aidx in assoc_indices:
                sched = assoc_schedules.get(aidx, {}).get(team_huddle_name)
                if sched and (sched >= int1_start and sched < int1_end):
                    first_interval_count += 1
            percent = (first_interval_count/total)*100
            team_huddle_stats.append({'Shift_start': ss_key, 'Total': total, 'First_interval_count': first_interval_count, 'Percent_first_interval': percent})

team_huddle_df = pd.DataFrame(team_huddle_stats)

# ---------- Logging to console ----------
print("\n==== Scheduling Summary ====")
print("Total associates processed:", len(assoc_df))
print("Total scheduled associate-meetings (count):", sum(len(v) for v in assoc_schedules.values()))
print("Total unscheduled entries:", len(unscheduled))
if len(unscheduled) > 0:
    print("\nUnscheduled (sample up to 10):")
    for u in unscheduled[:10]:
        print(u)
print("\nTeam Huddle distribution (per shift):")
if not team_huddle_df.empty:
    print(team_huddle_df.to_string(index=False))
else:
    print("No Team_Huddle statistics available or no Team_Huddle meetings scheduled.")

# ---------- Write output workbook ----------
print("\nWriting outputs to:", OUTPUT_XLSX)
# Build a dict of DataFrames to write
with pd.ExcelWriter(OUTPUT_XLSX, engine='openpyxl') as writer:
    # Write modified Associate_Roster
    assoc_df.to_excel(writer, sheet_name='Associate_Roster', index=False)
    # Write modified Manager_Roster
    manager_df.to_excel(writer, sheet_name='Manager_Roster', index=False)
    # Write updated Schedule_heatmap
    heatmap_df.to_excel(writer, sheet_name='Schedule_heatmap', index=False)
    # Write Constraint and Summary
    constraint_df.to_excel(writer, sheet_name='Constraint', index=False)
    summary_df.to_excel(writer, sheet_name='Summary_Associate_Meetings', index=False)
    pd.DataFrame(unscheduled).to_excel(writer, sheet_name='Unscheduled', index=False)
    if not team_huddle_df.empty:
        team_huddle_df.to_excel(writer, sheet_name='Team_Huddle_Distribution', index=False)

print("Done. Output written.")

# End of script
