import pandas as pd
import psycopg
import os


DATABASE_URL = "postgresql://postgres:Aanchal%40086@db.gcugsrwshbhoqkmyrapc.supabase.co:5432/postgres"


def import_csv(csv_file):

    table_name = os.path.splitext(
        os.path.basename(csv_file)
    )[0].lower()

    df = pd.read_csv(csv_file)

    if df.empty:
        print(f"{csv_file} is empty")
        return

    # Clean column names
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace(r"[^a-zA-Z0-9_]", "", regex=True)
    )

    columns = list(df.columns)

    column_sql = ", ".join(
        f'"{col}"' for col in columns
    )

    placeholders = ", ".join(
        ["%s"] * len(columns)
    )

    insert_sql = f"""
    INSERT INTO "{table_name}"
    ({column_sql})
    VALUES ({placeholders})
    """

    print("\n---------------------------------")
    print(f"Table   : {table_name}")
    print(f"Rows    : {len(df)}")
    print(f"Columns : {columns}")

    conn = None

    try:

        conn = psycopg.connect(
            DATABASE_URL,
            connect_timeout=30
        )

        with conn.cursor() as cur:

            cur.execute("SELECT 1")
            print("Database connection OK")

            rows = [
                tuple(
                    None if pd.isna(v) else v
                    for v in row
                )
                for row in df.itertuples(
                    index=False,
                    name=None
                )
            ]

            cur.executemany(
                insert_sql,
                rows
            )

        conn.commit()

        print(
            f"SUCCESS: {len(rows)} rows loaded into {table_name}"
        )

    except Exception as e:

        if conn:
            conn.rollback()

        print(f"FAILED: {table_name}")
        print(type(e).__name__)
        print(e)

    finally:

        if conn:
            conn.close()


if __name__ == "__main__":

    import_csv(
        r"C:\Users\91931\2w_dashboard\fact_sales_monthly.csv"
    )