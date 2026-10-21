"""
Groop Groop cell roster generator.

Reads member availability (Google Form responses) and the committee's session
schedule from one Google Sheet, then writes a "Roster" tab and a "Load Summary"
tab back to the same sheet.

Rules
  Hard: a member is never rostered on one of their block-out dates.
  Hard: a member is only rostered for roles they ticked on the form.
  Hard: no one does two roles in the same week.
  Soft: no one is rostered within BIRTHDAY_BUFFER_DAYS of their birthday
        (relaxed only if nobody else can fill the slot, and flagged).
  Balancing: effort points (Word = 2, others = 1) are spread so each person's
        load is proportional to their preferred frequency, and nobody is
        rostered again before their preferred gap has passed if avoidable.

Usage
  python roster.py                 # read sheet, write Roster + Load Summary tabs
  python roster.py --dry-run       # read sheet, print roster, write nothing
  python roster.py --local responses.csv schedule.csv   # offline test with CSV exports
"""

import argparse
import csv
import datetime as dt
import random
import re
import sys
from dataclasses import dataclass, field

# ---------------------------------------------------------------- config ----
SPREADSHEET_ID = "PASTE_YOUR_SHEET_ID_HERE"  # the long id in the sheet's URL
CREDENTIALS_FILE = "service_account.json"
RESPONSES_TAB = "Form Responses 1"
SCHEDULE_TAB = "Schedule"
ROSTER_TAB = "Roster"
SUMMARY_TAB = "Load Summary"

ROLES = ["Prayer", "Word", "Worship (Vocals)", "Worship (Instru)"]
EFFORT = {"Prayer": 1, "Word": 2, "Worship (Vocals)": 1, "Worship (Instru)": 1}
BIRTHDAY_BUFFER_DAYS = 3  # +/- days around the birthday treated as "birthday weekend"
DAYS_PER_MONTH = 30.4
RANDOM_SEED = 42  # fixed so re-runs give the same roster; change to reshuffle ties

DATE_FORMATS = ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y", "%d %b %Y", "%d %B %Y")
TRUTHY = {"y", "yes", "true", "x", "1", "✓", "✔"}
UNFILLED = "⚠ UNFILLED"


# ----------------------------------------------------------------- model ----
@dataclass
class Member:
    name: str
    birthday: dt.date | None
    roles: set[str]
    freq_months: float
    blockouts: set[dt.date]
    other_serving: str = ""
    comments: str = ""


@dataclass
class Week:
    date: dt.date
    roles: list[str]
    note: str = ""


@dataclass
class Slot:
    person: str
    flags: list[str] = field(default_factory=list)


# --------------------------------------------------------------- parsing ----
def parse_date(text: str) -> dt.date | None:
    text = (text or "").strip()
    if not text:
        return None
    for fmt in DATE_FORMATS:
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    print(f"  ! could not parse date '{text}' - ignored", file=sys.stderr)
    return None


def parse_freq(text: str) -> float:
    m = re.search(r"\d+(\.\d+)?", text or "")
    return float(m.group()) if m else 2.0  # sensible default if left blank


def find_col(headers: list[str], *keywords: str) -> int | None:
    for i, h in enumerate(headers):
        if any(k in h.lower() for k in keywords):
            return i
    return None


def parse_members(rows: list[list[str]]) -> list[Member]:
    headers, data = rows[0], rows[1:]
    c_name = find_col(headers, "name")
    c_bday = find_col(headers, "birthday")
    c_roles = find_col(headers, "weekly", "cell sessions")
    c_other = find_col(headers, "other serving")
    c_freq = find_col(headers, "frequency")
    c_comm = find_col(headers, "comment")
    c_blocks = [i for i, h in enumerate(headers) if "block" in h.lower()]
    if c_name is None or c_roles is None:
        sys.exit("Could not find the name / serving-area columns in the responses tab.")

    def cell(row, i):
        return row[i].strip() if i is not None and i < len(row) else ""

    members: dict[str, Member] = {}
    for row in data:  # later submissions overwrite earlier ones from the same person
        name = cell(row, c_name)
        if not name:
            continue
        ticked = cell(row, c_roles)
        roles = {r for r in ROLES if r.lower() in ticked.lower()}
        members[name.lower()] = Member(
            name=name,
            birthday=parse_date(cell(row, c_bday)),
            roles=roles,
            freq_months=parse_freq(cell(row, c_freq)),
            blockouts={d for d in (parse_date(cell(row, i)) for i in c_blocks) if d},
            other_serving=cell(row, c_other),
            comments=cell(row, c_comm),
        )
    return list(members.values())


def parse_schedule(rows: list[list[str]]) -> list[Week]:
    headers, data = rows[0], rows[1:]
    c_date = find_col(headers, "date")
    c_note = find_col(headers, "note")
    role_cols = {r: i for r in ROLES for i, h in enumerate(headers) if h.strip().lower() == r.lower()}
    missing = set(ROLES) - role_cols.keys()
    if c_date is None or missing:
        sys.exit(f"Schedule tab needs a Date column and columns {ROLES}. Missing: {missing or 'Date'}")

    weeks = []
    for row in data:
        d = parse_date(row[c_date] if c_date < len(row) else "")
        if not d:
            continue
        needed = [r for r, i in role_cols.items() if i < len(row) and row[i].strip().lower() in TRUTHY]
        note = row[c_note].strip() if c_note is not None and c_note < len(row) else ""
        weeks.append(Week(d, needed, note))
    return sorted(weeks, key=lambda w: w.date)


# ------------------------------------------------------------- scheduler ----
def near_birthday(m: Member, day: dt.date) -> bool:
    if not m.birthday:
        return False
    for year in (day.year - 1, day.year, day.year + 1):
        try:
            bday = m.birthday.replace(year=year)
        except ValueError:  # 29 Feb in a non-leap year
            bday = dt.date(year, 2, 28)
        if abs((day - bday).days) <= BIRTHDAY_BUFFER_DAYS:
            return True
    return False


def can_fill(roles: list[str], members: list[Member], day: dt.date, busy: set[str]) -> bool:
    """True if every role can get a distinct, available person (birthdays ignored)."""
    if not roles:
        return True
    role, rest = roles[0], roles[1:]
    return any(can_fill(rest, members, day, busy | {m.name})
               for m in members
               if role in m.roles and day not in m.blockouts and m.name not in busy)


def build_roster(members: list[Member], weeks: list[Week]) -> tuple[dict, dict, dict]:
    rng = random.Random(RANDOM_SEED)
    tiebreak = {m.name: rng.random() for m in members}
    load = {m.name: 0.0 for m in members}
    last = {m.name: None for m in members}
    counts = {m.name: {r: 0 for r in ROLES} for m in members}
    roster: dict[dt.date, dict[str, Slot]] = {}

    for wk in weeks:
        roster[wk.date] = {}
        busy: set[str] = set()

        def eligible(role):
            return [m for m in members if role in m.roles and wk.date not in m.blockouts]

        # fill the hardest-to-staff role first, Word before others on ties
        for role in sorted(wk.roles, key=lambda r: (len(eligible(r)), -EFFORT[r])):
            pool = [m for m in eligible(role) if m.name not in busy]
            tiers = [  # relax the birthday rule only if nobody else can serve
                ([m for m in pool if not near_birthday(m, wk.date)], []),
                (pool, ["birthday week - no alternative"]),
            ]
            cands, flags = next(((c, f) for c, f in tiers if c), ([], []))
            if not cands:
                roster[wk.date][role] = Slot(UNFILLED, ["everyone blocked out / already serving this week"])
                continue

            def score(m: Member):
                gap_days = (wk.date - last[m.name]).days if last[m.name] else 10_000
                too_soon = gap_days < m.freq_months * DAYS_PER_MONTH
                # load scaled by preferred frequency: a "monthly" person should
                # carry ~2x the load of a "2-monthly" person
                normalized = (load[m.name] + EFFORT[role]) * m.freq_months
                return (too_soon, normalized, -gap_days, tiebreak[m.name])

            remaining = [r for r in wk.roles if r not in roster[wk.date] and r != role]
            fillable_now = can_fill(remaining, members, wk.date, busy)
            ranked = sorted(cands, key=score)
            # take the best-scoring person whose pick still leaves the rest of
            # the week fillable (no one can double up, so choice order matters)
            pick = next((m for m in ranked
                         if not fillable_now or can_fill(remaining, members, wk.date, busy | {m.name})),
                        ranked[0])
            roster[wk.date][role] = Slot(pick.name, list(flags))
            busy.add(pick.name)
            load[pick.name] += EFFORT[role]
            last[pick.name] = wk.date
            counts[pick.name][role] += 1

    return roster, load, counts


# ---------------------------------------------------------------- output ----
def roster_table(weeks, roster) -> list[list[str]]:
    out = [["Date", *ROLES, "Notes"]]
    for wk in weeks:
        if not wk.roles:
            out.append([wk.date.strftime("%d/%m/%Y"), *[""] * len(ROLES), wk.note or "No session"])
            continue
        notes = [wk.note] if wk.note else []
        row = [wk.date.strftime("%d/%m/%Y")]
        for r in ROLES:
            slot = roster[wk.date].get(r)
            row.append(slot.person if slot else "-")
            if slot and slot.flags:
                notes.append(f"{r}: {', '.join(slot.flags)}")
        out.append(row + ["; ".join(notes)])
    return out


def summary_table(members, load, counts, weeks) -> list[list[str]]:
    if weeks:
        span_months = max(1.0, (weeks[-1].date - weeks[0].date).days / DAYS_PER_MONTH)
    else:
        span_months = 1.0
    out = [["Name", "Preferred freq (months)", *ROLES, "Effort points",
            "Target (approx)", "Other serving", "Comments"]]
    total_effort = sum(load.values())
    total_share = sum(1 / m.freq_months for m in members if m.roles) or 1
    for m in sorted(members, key=lambda m: -load[m.name]):
        target = total_effort * (1 / m.freq_months) / total_share if m.roles else 0
        out.append([m.name, m.freq_months, *[counts[m.name][r] for r in ROLES],
                    load[m.name], round(target, 1), m.other_serving, m.comments])
    out.append([])
    out.append([f"Roster covers ~{span_months:.1f} months; Word = 2 points, other roles = 1."])
    return out


# ----------------------------------------------------------------- I/O ------
def open_sheet():
    import gspread
    gc = gspread.service_account(filename=CREDENTIALS_FILE)
    return gc.open_by_key(SPREADSHEET_ID)


def write_tab(sh, title, rows):
    import gspread
    try:
        ws = sh.worksheet(title)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=max(100, len(rows) + 10), cols=12)
    ws.update(range_name="A1", values=[[str(c) for c in r] for r in rows])
    ws.format("A1:Z1", {"textFormat": {"bold": True}})
    ws.freeze(rows=1)


def read_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.reader(f))


def main():
    ap = argparse.ArgumentParser(description="Generate the Groop Groop cell roster.")
    ap.add_argument("--dry-run", action="store_true", help="print only, don't write to the sheet")
    ap.add_argument("--local", nargs=2, metavar=("RESPONSES_CSV", "SCHEDULE_CSV"),
                    help="read from CSV files instead of Google Sheets (implies --dry-run)")
    args = ap.parse_args()

    if args.local:
        resp_rows, sched_rows = read_csv(args.local[0]), read_csv(args.local[1])
        sh = None
    else:
        sh = open_sheet()
        resp_rows = sh.worksheet(RESPONSES_TAB).get_all_values()
        sched_rows = sh.worksheet(SCHEDULE_TAB).get_all_values()

    members = parse_members(resp_rows)
    weeks = parse_schedule(sched_rows)
    print(f"{len(members)} members, {len(weeks)} weeks")

    roster, load, counts = build_roster(members, weeks)
    r_rows = roster_table(weeks, roster)
    s_rows = summary_table(members, load, counts, weeks)

    if args.dry_run or args.local:
        for r in r_rows:
            print(" | ".join(str(c) for c in r))
        print()
        for r in s_rows:
            print(" | ".join(str(c) for c in r[:7]))
        return

    write_tab(sh, ROSTER_TAB, r_rows)
    write_tab(sh, SUMMARY_TAB, s_rows)
    print(f"Wrote '{ROSTER_TAB}' and '{SUMMARY_TAB}' tabs.")


if __name__ == "__main__":
    main()
