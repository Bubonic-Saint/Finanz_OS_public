import pandas as pd


def transactions_to_table(data: dict) -> pd.DataFrame:
    df = pd.DataFrame(data["transactions"])[["date", "merchant", "amount", "currency", "description"]]
    df["description"] = (
        df["description"]
        .str.extract(r"remittanceinformation:(.+)", expand=False)
        .str.split("KAUFUMSATZ").str[0]
        .str.replace(r"^NR\s+\w+\s+\d+\s+", "", regex=True)
        .str.strip()
    )
    return df
