# poolroute: weekly pool-route optimizer (no AI tokens needed)

## One-time setup
    pip install -r requirements.txt

## Every week
    python poolroute.py

Drop the new weekly schedule (any name starting with `SCHEDULE`, e.g. `SCHEDULE_10.11.2026_-10.17.2026.xlsx`)
and the contacts file (any name starting with `CONTACTS`) in this folder. The newest of each is used automatically.
The depot defaults to 2850 Glades Circle, Bay #4, Weston, FL 33327; override with `--depot "address"` or `--depot none`.
You can still pass paths explicitly: `python poolroute.py SCHED.xlsx CONTACTS.xls`.

Output: `optimized_routes_YYYYMMDD.xlsx`
- **Summary**: per tech/day: current miles & drive time vs. optimized (both versions)
- **Shortest Miles** / **Fastest Time**: stop-by-stop order, address, gate/notes, leg miles/min, running clock
- **... (maps)**: one Google Maps link per tech/day to hand to the tech
- **Check These**: stops with no address, or fuzzy name matches to verify

Options: `--service-min 25` (minutes per pool), `--only MIKA`, `--day TUE`, `--out file.xlsx`.
`--depot none` optimizes each day with a free start/end point.

## Notes
- Needs internet the first time (Census + OpenStreetMap geocoding, OSRM road distances). Addresses are cached in `cache/`, so later weeks only look up NEW addresses.
- Reads the "Ativos" (active) sheet of the contacts file; the Excel read-only "encryption" is handled automatically.
- Stops are re-ordered within each tech's day. It does not move stops between techs or days.
- `--dry-run` uses FAKE coordinates, for testing only.
