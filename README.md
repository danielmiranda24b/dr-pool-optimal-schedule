# poolroute: weekly pool-route optimizer (no AI tokens needed)

## One-time setup
    pip install -r requirements.txt

## Every week
    python poolroute.py SCHEDULE.xlsx CONTACTS.xls --depot "your warehouse/home address"

Output: `optimized_routes_YYYYMMDD.xlsx`
- **Summary**: per tech/day: current miles & drive time vs. optimized (both versions)
- **Shortest Miles** / **Fastest Time**: stop-by-stop order, address, gate/notes, leg miles/min, running clock
- **... (maps)**: one Google Maps link per tech/day to hand to the tech
- **Check These**: stops with no address, or fuzzy name matches to verify

Options: `--service-min 25` (minutes per pool), `--only MIKA`, `--day TUE`, `--out file.xlsx`.
Omit `--depot` and each day is optimized with a free start/end point.

## Notes
- Needs internet the first time (Census + OpenStreetMap geocoding, OSRM road distances). Addresses are cached in `cache/`, so later weeks only look up NEW addresses.
- Reads the "Ativos" (active) sheet of the contacts file; the Excel read-only "encryption" is handled automatically.
- Stops are re-ordered within each tech's day. It does not move stops between techs or days.
- `--dry-run` uses FAKE coordinates, for testing only.
