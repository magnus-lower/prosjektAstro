from skyfield.api import load, Topos
from skyfield.almanac import risings_and_settings
from skyfield.positionlib import Geocentric
from pytz import timezone
from timezonefinder import TimezoneFinder
import pandas as pd
from datetime import datetime
import numpy as np

# ---- Konfigurer fødselsdata ----
birth_data = {
    'year': 1990,
    'month': 1,
    'day': 1,
    'hour': 12,
    'minute': 0,
    'latitude': 59.9139,   # Oslo
    'longitude': 10.7522
}

# ---- Last inn efemerider og tidsskala ----
eph = load('de406.bsp')
ts = load.timescale()

# ---- Finn tidssone for fødested ----
tf = TimezoneFinder()
tz_name = tf.timezone_at(lng=birth_data['longitude'], lat=birth_data['latitude'])
tz = timezone(tz_name)

local_dt = tz.localize(datetime(birth_data['year'], birth_data['month'], birth_data['day'],
                                birth_data['hour'], birth_data['minute']))
utc_dt = local_dt.astimezone(timezone('UTC'))
t = ts.utc(utc_dt.year, utc_dt.month, utc_dt.day, utc_dt.hour, utc_dt.minute)

# ---- Lag observatør ----
observer = eph['earth'] + Topos(latitude_degrees=birth_data['latitude'],
                                longitude_degrees=birth_data['longitude'])

# ---- Zodiac-segmentering ----
def zodiac_segment(deg):
    signs = ['ari', 'tau', 'gem', 'can', 'leo', 'vir', 'lib', 'sco', 'sag', 'cap', 'aqu', 'pis']
    sign = signs[int(deg // 30) % 12]
    segment = int((deg % 30) // 7.5)
    suffix = 'abcd'[segment]
    return f"{sign}{suffix}"

# ---- Husberegning med 4-deling ----
def house_segment(lon, houses):
    for i in range(12):
        start = houses[i]
        end = houses[(i + 1) % 12]
        # Sørg for kontinuitet rundt 0/360
        if end < start:
            end += 360
        adj_lon = lon if lon >= start else lon + 360
        if start <= adj_lon < end:
            segment_size = (end - start) / 4
            pos_in_segment = int((adj_lon - start) // segment_size)
            return f"{i+1}{'abcd'[pos_in_segment]}"
    return "??"

# ---- Hus og Asc/MC via Placidus (forenklet her med ecliptic_longitude) ----
from skyfield.api import wgs84

# Hent ekliptisk lengde for husspissene (for enkelhet: dummy 12 hus jevnt fordelt fra ASC)
astrometric = observer.at(t)
ecl = astrometric.observe(eph['sun']).apparent().ecliptic_latlon()[0].degrees
asc = (ecl + 90) % 360
mc = (ecl + 0) % 360
houses = [(asc + i*30) % 360 for i in range(12)]  # jevnfordelt forenkling

# ---- Planeter å sjekke ----
planet_keys = {
    'Sun': 'sun',
    'Moon': 'moon',
    'Mercury': 'mercury',
    'Venus': 'venus',
    'Mars': 'mars',
    'Jupiter': 'jupiter barycenter',
    'Saturn': 'saturn barycenter',
    'Uranus': 'uranus barycenter',
    'Neptune': 'neptune barycenter',
    'Pluto': 'pluto barycenter'
}

# ---- Beregn posisjoner ----
planet_segments = []
for name, key in planet_keys.items():
    astrometric = observer.at(t).observe(eph[key]).apparent()
    lon = astrometric.ecliptic_latlon()[0].degrees % 360
    z_seg = zodiac_segment(lon)
    h_seg = house_segment(lon, houses)
    planet_segments.extend([z_seg, h_seg, ""])

# ---- Sett sammen hele raden ----
asc_seg = zodiac_segment(asc)
mc_seg = zodiac_segment(mc)

row = [asc_seg, "", mc_seg, ""] + planet_segments

# ---- Lag DataFrame og skriv til Excel ----
df = pd.DataFrame([row])
df.to_excel("astrologisk_rad.xlsx", index=False, header=False)

print("Ferdig! Resultat lagret i 'astrologisk_rad.xlsx'")