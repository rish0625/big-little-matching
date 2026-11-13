"""Generate ranked big/little pairing suggestions from preference matrices."""

import csv
import random
import re
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


SUGGESTION_COUNT = 10
# Optional: enter one or two big names here to require them to take twins.
# If one name is provided, the second is selected automatically.
# Leave empty to select the two most-requested bigs automatically.
TWIN_BIG_NAMES = []


def clean(value):
    return " ".join(value.strip().split())


def split_names(value):
    """Return names from a cell, including the two names in (A)/(B) cells."""
    value = clean(value)
    if not value:
        return []
    marked = re.findall(r"\([ab]\)\s*([^()]+?)(?=\s*\([ab]\)|$)", value, re.I)
    if marked:
        return [clean(name) for name in marked if clean(name)]
    # Some cells omit parentheses around the A/B labels.
    marked = re.findall(r"(?:^|\s)[ab]\)\s*(.*?)(?=\s+[ab]\)\s*|$)", value, re.I)
    if marked:
        return [clean(name) for name in marked if clean(name)]
    return [value]


def matrix_names(path):
    with open(path, newline="", encoding="utf-8-sig") as source:
        rows = list(csv.reader(source))
    header_index = next((i for i, row in enumerate(rows)
                         if row and not clean(row[0]) and sum(bool(clean(x)) for x in row[1:]) > 1), None)
    if header_index is None:
        raise ValueError(f"{path}: expected a ranking matrix with names in a header row")
    names = [clean(name) for name in rows[header_index][1:] if clean(name)]
    return names, rows, header_index


def read_matrix(path, choice_names):
    names, rows, header_index = matrix_names(path)
    lookup = {name.casefold(): name for name in choice_names}
    ranks = {name: {} for name in names}
    for row in rows[header_index + 1:]:
        if len(row) < 2:
            continue
        try:
            rank = int(clean(row[0]))
        except ValueError:
            continue
        for column, name in enumerate(names, start=1):
            if column >= len(row):
                continue
            for choice in split_names(row[column]):
                canonical = lookup.get(choice.casefold())
                if canonical and canonical not in ranks[name]:
                    ranks[name][canonical] = rank
    return names, ranks


def hungarian(cost):
    """Minimum-cost square assignment; returns row-to-column indices."""
    n = len(cost)
    u, v, p, way = [0] * (n + 1), [0] * (n + 1), [0] * (n + 1), [0] * (n + 1)
    for i in range(1, n + 1):
        p[0] = i
        j0 = 0
        minv = [float("inf")] * (n + 1)
        used = [False] * (n + 1)
        while True:
            used[j0] = True
            i0, delta, j1 = p[j0], float("inf"), 0
            for j in range(1, n + 1):
                if not used[j]:
                    cur = cost[i0 - 1][j - 1] - u[i0] - v[j]
                    if cur < minv[j]:
                        minv[j], way[j] = cur, j0
                    if minv[j] < delta:
                        delta, j1 = minv[j], j
            for j in range(n + 1):
                if used[j]:
                    u[p[j]] += delta
                    v[j] -= delta
                else:
                    minv[j] -= delta
            j0 = j1
            if p[j0] == 0:
                break
        while True:
            j1 = way[j0]
            p[j0] = p[j1]
            j0 = j1
            if j0 == 0:
                break
    result = [0] * n
    for j in range(1, n + 1):
        result[p[j] - 1] = j - 1
    return result


def generate(little_names, little_ranks, big_names, big_ranks, count=SUGGESTION_COUNT,
             requested_twins=None):
    if len(little_names) != len(big_names) + 2:
        raise ValueError(
            f"This setup expects exactly two extra littles for twins; found "
            f"{len(little_names)} littles and {len(big_names)} bigs."
        )
    requested = {big: 0 for big in big_names}
    for prefs in little_ranks.values():
        for big in prefs:
            if big in requested:
                requested[big] += 1

    # Each candidate has capacity two; every other big has capacity one.
    candidate_order = sorted(big_names, key=lambda big: (-requested[big], big.casefold()))
    if requested_twins:
        lookup = {name.casefold(): name for name in big_names}
        selected = [lookup.get(name.strip().casefold()) for name in requested_twins]
        if len(selected) not in (1, 2) or None in selected:
            raise ValueError("TWIN_BIG_NAMES must contain one or two names from the big preferences CSV.")
        if len(set(selected)) != len(selected):
            raise ValueError("TWIN_BIG_NAMES must contain different bigs.")
        if len(selected) == 1:
            second = next(big for big in candidate_order if big != selected[0])
            selected.append(second)
        twin_combinations = [tuple(selected)]
    else:
        twin_combinations = [(candidate_order[i], candidate_order[j])
                             for i in range(len(candidate_order))
                             for j in range(i + 1, len(candidate_order))]
        twin_combinations.sort(key=lambda pair: (-(requested[pair[0]] + requested[pair[1]]),
                                                 pair[0].casefold(), pair[1].casefold()))
    suggestions, seen = [], set()
    rng = random.Random(20261113)
    for offset in range(max(count * 8, len(twin_combinations))):
        twins = twin_combinations[offset % len(twin_combinations)]
        slots = big_names + list(twins)
        rng.shuffle(slots)
        scale = 100
        cost = []
        for little in little_names:
            row = []
            for big in slots:
                # Unranked roster members go after all ranked choices.
                lr = little_ranks[little].get(big, len(big_names) + 1)
                br = big_ranks[big].get(little, len(little_names) + 1)
                row.append((lr + br) * scale + rng.randrange(scale))
            cost.append(row)
        assignment = hungarian(cost)
        pairs = tuple(sorted(((slots[col], little_names[i]) for i, col in enumerate(assignment)),
                             key=lambda pair: (pair[0].casefold(), pair[1].casefold())))
        if pairs not in seen:
            seen.add(pairs)
            suggestions.append((twins, pairs))
            if len(suggestions) == count:
                break
    return suggestions


def write_output(path, suggestions, little_ranks, big_ranks):
    workbook = Workbook()
    workbook.remove(workbook.active)
    headers = ["Twin Big", "Big", "Suggested Little", "Big Rank", "Little Rank", "Ranked by Both?"]
    for number, (twins, pairs) in enumerate(suggestions, start=1):
        sheet = workbook.create_sheet(f"Suggestion {number}")
        sheet.append(headers)
        for big, little in pairs:
            lr = little_ranks[little].get(big, "Unranked")
            br = big_ranks[big].get(little, "Unranked")
            sheet.append([
                big if big in twins else "", big, little, br, lr,
                lr != "Unranked" and br != "Unranked",
            ])
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4472C4")
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column in sheet.columns:
            width = min(max(len(str(cell.value or "")) for cell in column) + 2, 32)
            sheet.column_dimensions[get_column_letter(column[0].column)].width = width
    workbook.save(path)


def main(args):
    if len(args) != 3:
        print("Usage: python biglittle.py [little preferences.csv] [big preferences.csv]")
        return 2
    little_names, _, _ = matrix_names(args[1])
    big_names, _, _ = matrix_names(args[2])
    little_names, little_ranks = read_matrix(args[1], big_names)
    big_names, big_ranks = read_matrix(args[2], little_names)
    suggestions = generate(little_names, little_ranks, big_names, big_ranks,
                           requested_twins=TWIN_BIG_NAMES)
    write_output("suggested pairings.xlsx", suggestions, little_ranks, big_ranks)
    print(f"Wrote {len(suggestions)} suggestions to suggested pairings.xlsx")
    if len(suggestions) < SUGGESTION_COUNT:
        print("Fewer than 10 unique pairings were found; review the generated suggestions.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv))
    except (OSError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        raise SystemExit(1)
