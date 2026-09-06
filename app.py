"""Read-only Streamlit view of savings_ledger.xlsx.

Run from the project folder, inside the conda env:

    streamlit run app.py
"""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

LEDGER = Path(__file__).with_name("savings_ledger.xlsx")


def payday_series(anchor: date, until: date) -> list[date]:
    if until < anchor:
        return []
    n = (until - anchor).days // 14
    return [anchor + timedelta(days=14 * i) for i in range(n + 1)]


@st.cache_data(ttl=30)
def load_ledger(path: str, mtime: float) -> tuple[dict, pd.DataFrame]:
    assumptions_df = pd.read_excel(path, sheet_name="Assumptions")
    assumptions = {
        str(row["key"]).strip(): row["value"]
        for _, row in assumptions_df.iterrows()
        if pd.notna(row.get("key"))
    }

    tx = pd.read_excel(path, sheet_name="Transactions")
    if tx.empty:
        tx = pd.DataFrame(
            columns=["date", "amount_transferred", "source", "notes"]
        )
    else:
        tx["date"] = pd.to_datetime(tx["date"]).dt.date
        tx["amount_transferred"] = pd.to_numeric(
            tx["amount_transferred"], errors="coerce"
        ).fillna(0.0)
        tx["source"] = tx.get("source", "scheduled").fillna("scheduled")
        tx["notes"] = tx.get("notes", "").fillna("")
        tx = tx.sort_values("date").reset_index(drop=True)

    starting = float(assumptions.get("starting_balance") or 0)
    tx["running_balance"] = starting + tx["amount_transferred"].cumsum()
    return assumptions, tx


def as_date(value) -> date:
    return pd.Timestamp(value).date()


def main() -> None:
    st.set_page_config(page_title="Savings expectation", layout="wide")
    st.title("Savings expectation")
    st.caption("Linear planned transfers. This page only reads the ledger.")

    if not LEDGER.exists():
        st.error(f"Missing {LEDGER.name}. Run savings_tracker.py first.")
        return

    assumptions, tx = load_ledger(str(LEDGER), LEDGER.stat().st_mtime)
    anchor = as_date(assumptions.get("payday_anchor"))
    planned = float(assumptions.get("biweekly_amount") or 0)
    starting = float(assumptions.get("starting_balance") or 0)
    today = date.today()

    posted = tx[tx["date"] <= today] if not tx.empty else tx
    current = float(posted["running_balance"].iloc[-1]) if len(posted) else starting

    upcoming = [d for d in payday_series(anchor, anchor + timedelta(days=3650)) if d > today]
    next_payday = upcoming[0] if upcoming else None

    left, mid, right, extra = st.columns(4)
    left.metric("Expected now", f"${current:,.2f}")
    mid.metric("Planned transfer", f"${planned:,.2f}")
    right.metric("Next payday", next_payday.isoformat() if next_payday else "—")
    extra.metric("Opening balance", f"${starting:,.2f}")

    st.divider()
    target = st.date_input("What should this be by?", value=date(today.year, 12, 31))
    what_if = st.slider(
        "What-if transfer amount (does not write to Excel)",
        min_value=0,
        max_value=max(int(planned * 3), 1000),
        value=int(planned),
        step=25,
    )

    future = [d for d in payday_series(anchor, target) if d > today]
    expected = current + what_if * len(future)
    st.subheader(f"${expected:,.2f}")
    st.write(
        f"{len(future)} planned deposits from tomorrow through {target.isoformat()} "
        f"at ${what_if:,.2f} each, on top of ${current:,.2f} already posted."
    )

    hist = posted.copy()
    if hist.empty:
        hist = pd.DataFrame({"date": [today], "running_balance": [current]})
    proj_rows = [{"date": today, "running_balance": current}]
    running = current
    for d in future:
        running += what_if
        proj_rows.append({"date": d, "running_balance": running})
    proj = pd.DataFrame(proj_rows)

    chart = pd.concat(
        [
            hist[["date", "running_balance"]].assign(series="posted"),
            proj.assign(series="projected"),
        ],
        ignore_index=True,
    )
    chart["date"] = pd.to_datetime(chart["date"])
    st.line_chart(chart, x="date", y="running_balance", color="series")

    st.subheader("Ledger")
    if tx.empty:
        st.info("No transactions yet. The first scheduled row lands on payday.")
    else:
        show = tx.copy()
        show["amount_transferred"] = show["amount_transferred"].map(lambda x: f"${x:,.2f}")
        show["running_balance"] = show["running_balance"].map(lambda x: f"${x:,.2f}")
        st.dataframe(show, hide_index=True, use_container_width=True)

    with st.expander("Assumptions (from Excel)"):
        st.write(
            {
                "payday_anchor": anchor.isoformat(),
                "biweekly_amount": planned,
                "starting_balance": starting,
            }
        )


if __name__ == "__main__":
    main()
