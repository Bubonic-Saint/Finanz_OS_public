import os
from datetime import date

import pandas as pd

from lunchFlow_Extraction import get_transactions
from transform import transactions_to_table
from state import read_last_run, write_last_run

ARCHIVES_DIR = os.path.join(os.path.dirname(__file__), '..', 'data', 'archives')
STAGING_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'data', 'pending_import.tsv'))


def write_staging(df):
    """Write transactions to a TSV staging file for VBA to pick up."""
    with open(STAGING_PATH, 'w', encoding='utf-8') as f:
        f.write("date\tamount\tmerchant\tdescription\n")
        for _, row in df.iterrows():
            desc = row.get("description", "")
            desc_str = str(desc) if pd.notna(desc) else ""
            merchant = str(row["merchant"]).replace("\t", " ")
            desc_str = desc_str.replace("\t", " ")
            f.write(f"{str(row['date'])[:10]}\t{float(row['amount'])}\t{merchant}\t{desc_str}\n")
    return len(df)


def main():
    today = date.today().isoformat()
    date_from = read_last_run()

    data = get_transactions(date_from=date_from)

    if not data["transactions"]:
        print("No new transactions since last run.")
        write_last_run(today)
        return

    df = transactions_to_table(data)

    os.makedirs(ARCHIVES_DIR, exist_ok=True)
    suffix = f"{date_from}_to_{today}" if date_from else today
    archive_path = os.path.join(ARCHIVES_DIR, f"transactions_{suffix}.csv")
    df.to_csv(archive_path, index=False, encoding="utf-8")

    count = write_staging(df)
    write_last_run(today)

    print(f"Staged {count} transaction(s) — open Excel to import.")
    print(df.to_string(index=False))
    print(f"\nArchive: {archive_path}")
    print(f"Staging: {STAGING_PATH}")


if __name__ == "__main__":
    main()
