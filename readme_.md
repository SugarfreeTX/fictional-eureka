### Useful extra commands already in the starter: 



python savings_tracker.py

python savings_tracker.py --add 250 "extra transfer"

python savings_tracker.py --project 2026-12-31

python savings_tracker.py --as-of 2026-10-02   # backfill / test

--project is the expectation tool: posted activity through today + remaining scheduled deposits until the target date. It does not pre-write 2027 rows into Excel.



### How to run Streamlit:

conda activate budget-manager

cd \~/projects/budget-manager

streamlit run app.py 

