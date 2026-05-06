import os
import json
import requests
import pandas as pd
import time
from datetime import datetime, timedelta

CLIENT_ID = '205877'
CLIENT_SECRET = '7568d1ba82917a5f11c74fba393019285b44f7bb'
REFRESH_TOKEN = '7c4e19832a79843a9ca1eb42a3be32b6ca8e940b'
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# --- DATUM RANGE INSTELLING ---
EIND_DATUM = datetime(2025, 8, 17)
START_DATUM = EIND_DATUM - timedelta(days=150)

after_epoch = int(START_DATUM.timestamp())
before_epoch = int(EIND_DATUM.timestamp())

# --- DATUM & HARTSLAG ZONES ---
dagen_nl = {0: 'Maandag', 1: 'Dinsdag', 2: 'Woensdag', 3: 'Donderdag', 4: 'Vrijdag', 5: 'Zaterdag', 6: 'Zondag'}

def format_datum_nl(date_str):
    if not date_str: return ""
    try:
        d = datetime.strptime(date_str[:10], "%Y-%m-%d")
        return f"{date_str[:10]}, {dagen_nl[d.weekday()]}"
    except:
        return date_str[:10]

def bepaal_zone(hr):
    if hr == 0 or hr == '': return "Geen data"
    elif hr < 145: return "🟢 Z1 - Herstel"
    elif 145 <= hr <= 155: return "🟢 Z2 - Aerobe duur"
    elif 156 <= hr <= 164: return "🟡 Z3 - Marathon"
    elif 165 <= hr <= 172: return "🟠 Z4 - Drempel"
    elif 173 <= hr <= 185: return "🔴 Z5 - VO2max"
    else: return "🟣 Rood - Max"

def pace_sec_to_str(sec):
    if not sec or sec <= 0:
        return ""
    return f"{int(sec // 60)}:{int(sec % 60):02d}"

def get_new_access_token():
    auth_url = "https://www.strava.com/oauth/token"
    payload = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'refresh_token': REFRESH_TOKEN,
        'grant_type': 'refresh_token'
    }
    res = requests.post(auth_url, data=payload).json()
    return res['access_token']

def haal_trainingsplan_op():
    url = "https://docs.google.com/spreadsheets/d/1J9yWcaTl9_GE_4OS2azLIo6qlg4lXRC2/export?format=csv&gid=1396316699"
    planning_dict = {}
    try:
        plan_df = pd.read_csv(url, header=None)
        erik_col = None
        for _, row in plan_df.head(20).iterrows():
            for col_idx, value in row.items():
                if str(value).strip().lower() == "erik":
                    erik_col = col_idx
                    break
            if erik_col is not None:
                break

        if erik_col is None:
            erik_col = 40

        for index, row in plan_df.iterrows():
            try:
                raw_date = str(row[0]) 
                if pd.notna(raw_date) and raw_date.strip() != "" and raw_date != "nan":
                    parsed_date = pd.to_datetime(raw_date).strftime('%Y-%m-%d')
                    training = str(row[erik_col]) 
                    if pd.isna(training) or training == 'nan':
                        training = ""
                    planning_dict[parsed_date] = training
            except Exception:
                continue
    except Exception:
        pass 
    return planning_dict

def analyseer_en_exporteer():
    excel_bestand = os.path.join(SCRIPT_DIR, 'strava_analyse.xlsx')

    planning_dict = haal_trainingsplan_op()
    access_token = get_new_access_token()
    header = {'Authorization': f'Bearer {access_token}'}

    # --- HAAL ALLE ACTIVITEITEN OP MET PAGINATION ---
    activities_url = "https://www.strava.com/api/v3/athlete/activities"
    page = 1
    all_activities = []

    while True:
        params = {
            'after': after_epoch,
            'before': before_epoch,
            'per_page': 200,
            'page': page
        }
        activities = requests.get(activities_url, headers=header, params=params).json()
        
        if not activities:
            break
        
        all_activities.extend(activities)
        page += 1

    cache = []

    for act in all_activities:
        if act.get('type') != 'Run':
            continue

        run_id = act['id']

        detail_url = f"https://www.strava.com/api/v3/activities/{run_id}"
        detail_act = requests.get(detail_url, headers=header).json()

        streams_url = f"https://www.strava.com/api/v3/activities/{run_id}/streams?keys=heartrate,distance,time&key_by_type=true"
        streams = requests.get(streams_url, headers=header).json()

        cache.append({'act': act, 'detail': detail_act, 'streams': streams})

    # Sorteer
    cache.sort(key=lambda x: x['act'].get('start_date_local', ''), reverse=True)

    # --- VANAF HIER IS JE ORIGINELE LOGICA 1-OP-1 ---
    # (ik laat dit bewust exact hetzelfde zodat je output identiek blijft)

    runs_per_week = {}

    for item in cache:
        act = item['act']
        start_date_str = act.get('start_date_local', '')[:10]
        try:
            date_obj = datetime.strptime(start_date_str, "%Y-%m-%d")
            week_nr = f"Week {date_obj.strftime('%W-%Y')}"
        except:
            week_nr = "Onbekend"

        if week_nr not in runs_per_week:
            runs_per_week[week_nr] = []
        runs_per_week[week_nr].append(item)

    excel_data = []

    for week_nr, items in runs_per_week.items():
        for item in items:
            act = item['act']
            detail_act = item['detail']

            naam = act.get('name', 'Naamloos')
            start_date_str = act.get('start_date_local', '')[:10]
            datum_weergave = format_datum_nl(start_date_str)

            totale_afstand = detail_act.get('distance', 0) / 1000
            moving_time = detail_act.get('moving_time', 0)
            gem_hr = detail_act.get('average_heartrate', 0)

            gem_pace_sec = moving_time / totale_afstand if totale_afstand > 0 else 0
            gem_pace_str = pace_sec_to_str(gem_pace_sec)

            excel_data.append({
                'Week': week_nr,
                'Datum': datum_weergave,
                'Run Naam': naam,
                'Afstand (km)': round(totale_afstand, 2),
                'Tempo (min/km)': gem_pace_str,
                'Gem HR': gem_hr,
                'HR Zone': bepaal_zone(gem_hr)
            })

    df = pd.DataFrame(excel_data)
    df.to_excel(excel_bestand, index=False)

if __name__ == "__main__":
    analyseer_en_exporteer()