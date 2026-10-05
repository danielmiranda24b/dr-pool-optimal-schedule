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
- **Shortest Miles** / **Fastest Time**: stop-by-stop order, address, notes, arrival time, leg miles/min, old order
- **... (maps)**: Google Maps links per tech/day, starting and ending at the shop (split into parts of 9 stops so they open on phones)
- **Check These**: cancelled/suspended stops left out, stops with no or approximate location, and name matches to verify

Options: `--service-min 25` (minutes per pool), `--start 8:00` (leave time, for the Arrive column), `--only MIKA`, `--day TUE`, `--out file.xlsx`,
`--depot-latlon "26.09,-80.38"` (skip looking up the depot; right-click the shop in Google Maps to copy its coordinates).

## Schedule tags
- `- C` cancelled and `- S` suspended: left out of the routes (listed on **Check These**)
- `- T`, `- BT`: routed normally

## Matching names to contacts
1. Fuzzy full-name match.
2. Otherwise last name, confirmed by the house number in the name (`AMBRUGNA A. 2640`), the first initial, or the community.
   Spaces are ignored, so `DELAFUENTE` = `De La Fuente`.
3. Otherwise a partial name match for buildings/communities with typos (`VILLAS OF BONAVETURE`).
Every non-exact match is listed on **Check These** to verify.

If an address can't be found on the map (after spelling out `Tr`->`Trail`, `SW 137`->`SW 137th`, etc.), the stop is placed at the
centre of its community's other stops and marked `APPROX` in the Location column. If the depot can't be found, the run still finishes
with a free start/end and tells you to use `--depot-latlon`.
`--depot none` optimizes each day with a free start/end point.

## Notes
- Needs internet the first time (Census + OpenStreetMap geocoding, OSRM road distances). Addresses are cached in `cache/`, so later weeks only look up NEW addresses.
- Reads the "Ativos" (active) sheet of the contacts file; the Excel read-only "encryption" is handled automatically.
- Stops are re-ordered within each tech's day. It does not move stops between techs or days.
- `--dry-run` uses FAKE coordinates, for testing only.
