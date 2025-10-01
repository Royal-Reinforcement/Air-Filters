import streamlit as st
import datetime as dt
import pandas as pd
import smartsheet
import smtplib

from email.message import EmailMessage
from collections import defaultdict

APP_NAME = 'Air Filter Scheduler'


@st.cache_data(ttl=300)
def smartsheet_to_dataframe(sheet_id):
    smartsheet_client = smartsheet.Smartsheet(st.secrets['smartsheet']['access_token'])
    sheet             = smartsheet_client.Sheets.get_sheet(sheet_id)
    columns           = [col.title for col in sheet.columns]
    rows              = []
    for row in sheet.rows: rows.append([cell.value for cell in row.cells])
    return pd.DataFrame(rows, columns=columns)


def first_full_week_monday(year, month):
    first_day       = dt.date(year, month, 1)
    first_monday    = first_day + dt.timedelta(days=(0 - first_day.weekday()) % 7)

    if (first_monday + dt.timedelta(days=4)).month != month: first_monday += dt.timedelta(days=7)
    return first_monday if first_monday.month == month else None


def month_weeks(year, month, max_weeks=4):
    start = first_full_week_monday(year, month)

    if not start: return []

    weeks = []
    for i in range(max_weeks):
        week_start = start + dt.timedelta(weeks=i)
        week_end = week_start + dt.timedelta(days=4)
        if week_start.month != month or week_end.month != month:
            break
        weeks.append((week_start, week_end))

    return weeks

def get_current_and_next_week(today):
    year            = today.year
    month           = today.month
    weeks           = month_weeks(year, month)
    current_num     = None
    current_range   = None
    next_num        = None
    next_range      = None

    for i, (week_start, week_end) in enumerate(weeks, 1):
        if week_start <= today <= week_end:
            current_num     = i
            current_range   = (week_start, week_end)
            break

    if current_num is None and weeks and today < weeks[0][0]:
        current_num     = 0
        current_range   = None
        next_num        = 1
        next_range      = weeks[0]

        return current_num, current_range, next_num, next_range

    if (not weeks) or (current_num is None and weeks and today > weeks[-1][1]) or (current_num == len(weeks)):
        next_year, next_month = (year + 1, 1) if month == 12 else (year, month + 1)

        next_weeks = month_weeks(next_year, next_month)
        
        if current_num == len(weeks):
            current_range   = weeks[-1]
            current_num     = len(weeks)

        next_num, next_range = (1, next_weeks[0]) if next_weeks else (None, None)

        return current_num, current_range, next_num, next_range

    if current_num:
        current_range   = weeks[current_num - 1]
        next_num        = current_num + 1
        next_range      = weeks[current_num]

    return current_num, current_range, next_num, next_range


def process_unit(group):
        arriving    = set(group["Start_Date"])
        departing   = set(group["Departure"])
        occupied    = set()

        for _, row in group.iterrows():
            range = pd.date_range(row["Start_Date"] + pd.Timedelta(days=1), row["Departure"] - pd.Timedelta(days=1))
            occupied.update(range)
        
        all_days    = set(pd.date_range(group["Start_Date"].min(), group["Departure"].max()))
        vacant      = all_days - occupied - arriving - departing
        
        return {
            "arriving": sorted(arriving),
            "departing": sorted(departing),
            "occupied": sorted(occupied),
            "vacant": sorted(vacant)
        }


def schedule_tasks(result, start, end, subset=None):
    schedule_dates      = pd.date_range(start, end)
    load                = defaultdict(int)
    assignments_per_day = defaultdict(list)

    units = subset if subset is not None else result.keys()

    for unit in units:
        if unit not in result:
            continue

        data = result[unit]

        arriving    = [d for d in data["arriving"]  if start <= d <= end]
        departing   = [d for d in data["departing"] if start <= d <= end]
        vacant      = [d for d in data["vacant"]    if start <= d <= end]
        occupied    = [d for d in data["occupied"]  if start <= d <= end]

        b2b         = sorted(set(arriving) & set(departing))

        arriving    = [d for d in arriving  if d not in b2b]
        departing   = [d for d in departing if d not in b2b]


        if vacant:
            candidates  = vacant
            status      = "VACANT"
        elif b2b:
            candidates  = b2b
            status      = "B2B"
        elif arriving:
            candidates  = arriving
            status      = "ARRIVAL"
        elif departing:
            candidates  = departing
            status      = "DEPARTURE"
        elif occupied:
            candidates  = occupied
            status      = "OCCUPIED"
        else:
            candidates  = list(schedule_dates)
            status      = "VACANT"

        chosen_day          = min(candidates, key=lambda d: load[d])
        load[chosen_day]    += 1

        assignments_per_day[chosen_day.strftime("%Y-%m-%d")].append({"unit": unit, "status": status})

    load_str = {day.strftime("%Y-%m-%d"): count for day, count in load.items()}

    return dict(assignments_per_day), load_str


def assignments_dict_to_df(assignments_per_day):
    records = []
    for date_str, units in assignments_per_day.items():
        for unit_info in units:
            records.append({
                "date": pd.to_datetime(date_str),
                "unit_code": unit_info["unit"],
                "status": unit_info["status"]
            })
    
    df = pd.DataFrame(records)
    df = df.sort_values(["date", "unit_code"]).reset_index(drop=True)
    return df


def email_dataframe_as_csv(df, filename, recipients, subject='Royal Reinforcement Notification', body='Please see the attached file.'):
        sender      = st.secrets['email']['username']
        password    = st.secrets['email']['password']
        msg         = EmailMessage()

        msg["From"]     = sender
        msg["To"]       = (', ').join(recipients) if isinstance(recipients, list) else recipients
        msg["Subject"]  = subject

        msg.set_content(body)
        msg.add_attachment(df.to_csv(index=False).encode("utf-8"), maintype='text', subtype='csv', filename=f'{filename}.csv')

        with smtplib.SMTP('smtp.office365.com',587) as s:
            s.starttls()
            s.login(sender, password)
            s.send_message(msg)


st.set_page_config(page_title=APP_NAME, page_icon='💨', layout='centered')

st.image(st.secrets['images']["rr_logo"], width=100)

st.title(APP_NAME)
st.info('Creation of an air filter schedule using the broad air filter changing cadence and occupancy data.')


current_year    = dt.datetime.now().year
next_year       = current_year + 1
report_url      = f"{st.secrets['escapia_1']}{current_year}{st.secrets['escapia_2']}{next_year}{st.secrets['escapia_3']}"

st.link_button('Download the **Housekeeping Report** from **Escapia**', url=report_url, type='secondary', use_container_width=True, help='Housekeeping Arrival Departure Report - Excel 1 line')

escapia_file = st.file_uploader(label='Housekeeping Arrival Departure Report - Excel 1 line.csv', type='csv')

if escapia_file is not None:

    df = pd.read_csv(escapia_file, index_col=False)
    df['Start_Date']    = pd.to_datetime(df['Start_Date'])
    df['Departure']     = pd.to_datetime(df['Departure'])
    df = df[['Unit_Code', 'Start_Date', 'Departure']]

    working_weeks = [1, 2, 3, 4]
    today = st.date_input('Date')
    current_week, current_range, next_week, next_range = get_current_and_next_week(today)
        
    result              = df.groupby("Unit_Code", group_keys=True).apply(process_unit, include_groups=False).to_dict()
    air_filter_schedule = smartsheet_to_dataframe(st.secrets['smartsheet']['sheets']['schedule'])

    if next_week in working_weeks:
        week_start  = pd.Timestamp(next_range[0])
        week_end    = pd.Timestamp(next_range[1])

        weekly_air_filters = air_filter_schedule[air_filter_schedule['Week'] == next_week]['Unit_Code'].tolist()
        assignments = schedule_tasks(result, week_start, week_end, subset=weekly_air_filters)

        with st.expander('Workload Distribution'):

            columns = st.columns(5)
            count = 0

            for date in sorted(assignments[1].keys()):
                columns[count].metric(pd.to_datetime(date).strftime('**%A**\n\n%m/%d/%Y'), assignments[1][date])
                count = (count + 1) % 5

            deliverable = assignments_dict_to_df(assignments[0])
            deliverable['date'] = deliverable['date'].dt.strftime('%Y-%m-%d')
            deliverable['day']  = pd.to_datetime(deliverable['date']).dt.strftime('%A')
            deliverable['week'] = next_week
            deliverable         = deliverable[['week', 'date', 'day', 'unit_code', 'status']]
            deliverable.columns = ['Week', 'Date', 'Day', 'Unit_Code', 'Status']
            deliverable = pd.merge(deliverable, air_filter_schedule[['Unit_Code','Ladder?','Filters','#']], on='Unit_Code', how='left')


        with st.expander('Schedule'):
            st.dataframe(deliverable, use_container_width=True, hide_index=True)

        
        if st.button('Send Schedule to Email', use_container_width=True, type='primary'):
            date1 = week_start.strftime('%m-%d')
            date2 = week_end.strftime('%m-%d-%Y')

            recipients = [
                st.secrets['recipients']['guest_services'],
                st.secrets['recipients']['inspectors'],
                st.secrets['recipients']['liaisons'],
                st.secrets['recipients']['maintenance']

            ]

            email_dataframe_as_csv(
                df=deliverable,
                filename=f'AFS_{date1}_{date2}',
                recipients=recipients,
                subject=f'Air Filter Schedule | {date1} - {date2}',
                body='Please see the attached Air Filter Schedule.'
                )
            
            st.toast(icon='📧', body='Email sent!')