from db_connect import get_connection


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

if __name__ == "__main__":
    conn = get_connection()
    countries = get_countries(conn)
    for country in countries:
        print(country)
    conn.close()
