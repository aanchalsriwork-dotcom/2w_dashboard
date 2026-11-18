from db_connect import get_connection

try:
    conn = get_connection()
    print("Connected successfully")
except:
    print("Database connection failed")


# Fetches country names from DB and returns a list of dictionaries with country code and name
def get_countries(conn):
    # Create a cursor object : tool used to send SQL commands to db
    cur = conn.cursor()
    # Execute runs the query and cursor now holds the result set.
    cur.execute("SELECT wb_code, country FROM country;")
    # Retrieve rows produced by the query from cursor object
    rows = (
        cur.fetchall()
    )  # returns a list of tuples, each tuple is a row in the result set
    cur.close()
    # Loop through the rows and create a list of dictionaries with column names as keys and row values as values
    return [{"code": r[0], "name": r[1]} for r in rows]


# Fetches indicator details from DB and returns a list of dictionaries with indicator code, database ID, source indicator ID, source, and frequency
def get_indicators(conn):
    cur = conn.cursor()
    cur.execute("""
        SELECT indicator_id, source, database_id, db_indicator_id, frequency, last_fetched_at
        FROM dim_indicator;
    """)
    rows = cur.fetchall()
    cur.close()
    return [
        {
            "indicator_id": r[0],
            "source": r[1],
            "database_id": r[2],
            "db_indicator_id": r[3],
            "frequency": r[4],
            "last_fetched_at": r[5],
        }
        for r in rows
    ]


def get_countries(conn):
    cur = conn.cursor()
    cur.execute("SELECT * FROM country;")
    rows = cur.fetchall()
    cur.close()
    countries = []

    for r in rows:
        country = {
            "region": r[0],
            "country": r[1],
            "wb_code": r[2],
            "currency_name": r[3],
            "currency_code": r[4],
        }

        countries.append(country)

    return countries
