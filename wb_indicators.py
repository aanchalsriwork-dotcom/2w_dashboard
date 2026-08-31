import requests
from pprint import pprint
from db_connect import get_connection
from fetchers import get_countries
from psycopg2.extras import execute_values
from datetime import date

# connection to the database
conn = get_connection()

# Get country codes
country_table = get_countries(conn)
country_codes = [country["wb_code"] for country in country_table]


API_URL = "https://data360api.worldbank.org/data360/data"

indicators = {
    "GDP": "IMF_WEO_NGDPD",  # Nominal GDP
    "GDP_GROWTH": "IMF_WEO_NGDP_RPCH",  # GDP Growth
    "GDP_PER_CAPITA": "IMF_WEO_NGDPPC",  # GDP per capita
    "UNEMPLOYMENT_RATE": "IMF_WEO_LUR",  # Unemployment rate
}


def fetch_data(indicator_code):
    params = {
        "DATABASE_ID": "IMF_WEO",
        "INDICATOR": indicators[indicator_code],
        "timePeriodFrom": 2025,
        "timePeriodTo": 2026,
    }

    response = requests.get(API_URL, params=params, timeout=30)
    response.raise_for_status()

    data = response.json()["value"]

    row = []
    for obs in data:
        wb_code = obs.get("REF_AREA")
        if wb_code not in country_codes:
            continue
        ind_code = indicator_code
        time_period = obs.get("TIME_PERIOD")
        value = obs.get("OBS_VALUE")
        row.append((wb_code, ind_code, time_period, value))

    return row


all_rows = []
for indicator in indicators.keys():
    obs = fetch_data(indicator)
    all_rows.extend(obs)

# Insert the fetched data into the database


# Convert the time_period to a date object
def period_to_date(time_period):
    return date(int(time_period), 1, 1)


# Define a function to insert/update rows in the database
def upsert_rows(conn, rows):
    # Create an empty list for cleaned rows
    clean_rows = []

    for row in rows:
        # Take the values from the row and give them names
        wb_code, ind_code, time_period, value = row

        # Check if the value is missing, skip row
        if value is None:
            continue  # skip this row, move to the next one

        # Convert the year into a date
        period_date = period_to_date(time_period)

        # Add the cleaned row to the list
        clean_rows.append((wb_code, ind_code, period_date, value))

    # Check if there are no rows to insert
    if not clean_rows:
        print("Nothing to write — no valid rows.")
        return

    sql = """
        INSERT INTO fact_macro (wb_code, ind_code, period_date, value)
        VALUES %s
        ON CONFLICT (wb_code, ind_code, period_date)
        DO UPDATE SET value = EXCLUDED.value, updated_at = now()
    """
    # Open a cursor to interact with the database
    with conn.cursor() as cur:
        # Execute the SQL query with the cleaned rows
        execute_values(cur, sql, clean_rows)
    # Save the changes to the database
    conn.commit()

    print(f"Saved {len(clean_rows)} rows to fact_macro")


# Run the function
upsert_rows(conn, all_rows)

# Close the database connection
conn.close()
