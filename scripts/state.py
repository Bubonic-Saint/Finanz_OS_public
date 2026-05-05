import os

LAST_RUN_FILE = os.path.join(os.path.dirname(__file__), '..', 'data', 'last_run.txt')


def read_last_run() -> str | None:
    if os.path.exists(LAST_RUN_FILE):
        with open(LAST_RUN_FILE) as f:
            return f.read().strip() or None
    return None


def write_last_run(run_date: str):
    with open(LAST_RUN_FILE, "w") as f:
        f.write(run_date)
