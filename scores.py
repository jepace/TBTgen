"""
Scores, seasons and standings (replaces the per-venue Numbers scoresheet).

Sheet tabs:
  Teams    team_id (TM####, permanent), name, venue_id (home venue), aliases (old names, comma-separated), active, notes
           Renaming a team keeps its history: results point at team_id, not the name.
  Results  date, venue_id, team_id, players, round1, round2, round3, tiebreak, photo, notes
           One row per team per night. tiebreak = total years off on the tiebreaker (only needed for ties).
           photo (on the winning team's row) = winners' picture under assets/ (pregame deck + social image).
  Seasons  venue_id, season, start_date, weeks, drop_lowest
           Seasons are per venue. A season holds the venue's first `weeks` game nights from start_date on.
Venues gain: team_max (default 6), size_penalty (points per extra member, default 1),
             penalty_per ("game": once per game, default; "round": every round).

Season rules (from the Numbers scoresheet):
  - Season total = sum of a team's game scores in the season.
  - Weighted score = total minus the team's `drop_lowest` lowest weeks, where a missed week counts as 0.
    In week 1 the weighted score is the total (otherwise everyone would be 0).
  - Ranking: weighted score, then season total, then number of weekly wins.
"""
import datetime as dt
from collections import defaultdict
from statistics import median

import build


def num(v, default=0):
    try:
        return float(v) if str(v).strip() not in ("", "None") else default
    except (TypeError, ValueError):
        return default


def clean_num(x):
    return int(x) if float(x).is_integer() else round(x, 2)


def venue_rules(venue):
    return {"team_max": int(num(venue.get("team_max"), 6) or 6),
            "size_penalty": num(venue.get("size_penalty"), 1),
            "penalty_per": "round" if str(venue.get("penalty_per") or "").strip().lower() == "round" else "game"}


def penalty(players, rules):
    """Points lost for members over the limit (per game, or every round)."""
    extra = max(0, int(num(players)) - rules["team_max"])
    pts = extra * rules["size_penalty"]
    return pts * 3 if rules["penalty_per"] == "round" else pts


def game_score(row, rules):
    rounds = [num(row.get(f"round{i}")) for i in (1, 2, 3)]
    pen = penalty(row.get("players"), rules)
    return {"rounds": rounds, "penalty": pen, "total": sum(rounds) - pen}


def teams(bank):
    return {str(r["team_id"]).strip(): r for r in bank.get("Teams", []) if str(r.get("team_id") or "").strip()}


def team_name(bank, tid):
    t = teams(bank).get(tid)
    return str(t["name"]) if t else tid


def find_team(bank, name, venue_id):
    """team_id for a name (current or old) at this venue, or None for a new team. Teams are per venue: two
    unrelated "Trivia Newton John" teams at different venues are different teams. A team that has played at this
    venue before (e.g. one that travels) also counts."""
    key = build.plain(name).strip().lower()
    played_here = {str(r.get("team_id")) for r in bank.get("Results", []) if r.get("venue_id") == venue_id}
    for tid, t in teams(bank).items():
        if t.get("venue_id") != venue_id and tid not in played_here:
            continue
        names = [str(t.get("name") or "")] + [a for a in str(t.get("aliases") or "").split(",")]
        if any(build.plain(n).strip().lower() == key for n in names if n.strip()):
            return tid
    return None


def team_label(bank, tid):
    """Team name, plus its venue when another team elsewhere has the same name."""
    ts = teams(bank)
    t = ts.get(tid)
    if not t:
        return tid
    name = str(t.get("name") or tid)
    if sum(1 for x in ts.values() if build.plain(x.get("name") or "").strip().lower() == build.plain(name).strip().lower()) > 1:
        v = bank["venues"].get(t.get("venue_id"), {})
        return f"{name} ({build.plain(v.get('name') or t.get('venue_id') or '?')})"
    return name


def night_results(bank, venue_id, date):
    """[(row, score dict)] for one night, best first (ties: lower tiebreak wins)."""
    venue = bank["venues"].get(venue_id, {})
    rules = venue_rules(venue)
    d = str(build.as_date(date))
    rows = [r for r in bank.get("Results", []) if r.get("venue_id") == venue_id and r.get("date")
            and str(build.as_date(r["date"])) == d]
    out = [(r, game_score(r, rules)) for r in rows]
    out.sort(key=lambda x: (-x[1]["total"], num(x[0].get("tiebreak"), 1e9)))
    return out


def seasons(bank, venue_id):
    return sorted((r for r in bank.get("Seasons", []) if r.get("venue_id") == venue_id and r.get("start_date")),
                  key=lambda r: str(build.as_date(r["start_date"])))


def season_for(bank, venue_id, date=None):
    """The venue's season that contains `date` (default: the latest season started by then)."""
    d = build.as_date(date) if date else dt.date.today()
    cur = None
    for s in seasons(bank, venue_id):
        if build.as_date(s["start_date"]) <= d:
            cur = s
    return cur


def season_standings(bank, venue_id, season=None, upto=None):
    """Standings for a venue's season (default: the current one), counting game nights up to `upto`."""
    s = season if isinstance(season, dict) else (
        next((x for x in seasons(bank, venue_id) if str(x.get("season")) == str(season)), None) if season
        else season_for(bank, venue_id, upto))
    venue = bank["venues"].get(venue_id, {})
    rules = venue_rules(venue)
    if not s:
        return {"season": None, "weeks": [], "rows": [], "venue_id": venue_id}
    start = build.as_date(s["start_date"])
    n_weeks = int(num(s.get("weeks"), 13) or 13)
    drop = int(num(s.get("drop_lowest"), 1))
    limit = build.as_date(upto) if upto else None
    dates = sorted({build.as_date(r["date"]) for r in bank.get("Results", [])
                    if r.get("venue_id") == venue_id and r.get("date") and build.as_date(r["date"]) >= start
                    and (limit is None or build.as_date(r["date"]) <= limit)})[:n_weeks]
    played = len(dates)
    by_team = defaultdict(dict)   # tid -> {date: total}
    winners = {}
    for d in dates:
        night = night_results(bank, venue_id, d)
        for r, sc in night:
            by_team[str(r["team_id"])][d] = sc["total"]
        if night:
            winners[d] = str(night[0][0]["team_id"])
    wins = defaultdict(int)
    for tid in winners.values():
        wins[tid] += 1
    rows = []
    for tid, weeks in by_team.items():
        all_weeks = [weeks.get(d, 0) for d in dates]      # a missed week counts as 0
        total = sum(all_weeks)
        dropped = sorted(all_weeks)[:drop] if played > 1 else []
        weighted = total - sum(dropped)
        rows.append({"team_id": tid, "team": team_name(bank, tid), "weeks_played": len(weeks), "total": clean_num(total),
                     "low": clean_num(min(all_weeks)) if all_weeks else 0, "weighted": clean_num(weighted),
                     "wins": wins[tid], "last": clean_num(weeks.get(dates[-1], 0)) if dates else 0})
    rows.sort(key=lambda r: (-r["weighted"], -r["total"], -r["wins"], r["team"].lower()))
    for i, r in enumerate(rows, 1):
        r["rank"] = i
    return {"venue_id": venue_id, "season": s.get("season"), "start": str(start), "weeks_total": n_weeks,
            "weeks_played": played, "drop_lowest": drop, "rules": rules,
            "weeks": [{"date": str(d), "winner_id": winners.get(d, ""), "winner": team_name(bank, winners.get(d, ""))
                       if winners.get(d) else ""} for d in dates],
            "rows": rows}


MEDALS = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣", "🔟"]


def standings_lines(st, weighted=True, limit=None):
    """'🥇 Extra Credit Crew (319)' lines, like the Numbers Output sheet."""
    key = "weighted" if weighted else "total"
    rows = sorted(st["rows"], key=lambda r: (-r[key], -r["total"], -r["wins"], r["team"].lower()))
    return [f"{MEDALS[i] if i < len(MEDALS) else f'{i + 1}.'} {r['team']} ({r[key]})"
            for i, r in enumerate(rows[:limit] if limit else rows)]


def demographics(bank, venue_id, since=None):
    """Teams and players per night, and the same for teams that came back (Numbers 'Demographics')."""
    nights = defaultdict(list)
    for r in bank.get("Results", []):
        if r.get("venue_id") == venue_id and r.get("date"):
            d = build.as_date(r["date"])
            if since is None or d >= build.as_date(since):
                nights[d].append(r)
    seen = defaultdict(int)
    for rows in nights.values():
        for r in rows:
            seen[str(r["team_id"])] += 1
    players = [num(r.get("players")) for rows in nights.values() for r in rows if num(r.get("players"))]
    repeat = [num(r.get("players")) for rows in nights.values() for r in rows
              if seen[str(r["team_id"])] > 1 and num(r.get("players"))]
    per_night = [len(rows) for rows in nights.values()]
    people = [sum(num(r.get("players")) for r in rows) for rows in nights.values()]
    rnd = lambda x: round(x, 1)
    return {"nights": len(nights), "teams_total": len(seen), "repeat_teams": sum(1 for v in seen.values() if v > 1),
            "teams_per_night": rnd(sum(per_night) / len(per_night)) if per_night else 0,
            "players_per_night": rnd(sum(people) / len(people)) if people else 0,
            "team_size_avg": rnd(sum(players) / len(players)) if players else 0,
            "team_size_median": rnd(median(players)) if players else 0,
            "repeat_team_size_avg": rnd(sum(repeat) / len(repeat)) if repeat else 0,
            "repeat_team_size_median": rnd(median(repeat)) if repeat else 0,
            "by_night": [{"date": str(d), "teams": len(nights[d]), "players": clean_num(sum(num(r.get("players")) for r in nights[d]))}
                         for d in sorted(nights)]}


def roster(bank, venue_id, date=None):
    """Teams to pre-fill a night's scorekeeper: the venue's teams that played this season or are based there."""
    st = season_standings(bank, venue_id, upto=date)
    ids = [r["team_id"] for r in st["rows"]]
    for tid, t in teams(bank).items():
        if t.get("venue_id") == venue_id and tid not in ids and not build.is_no(t.get("active")):
            ids.append(tid)
    return [{"team_id": tid, "name": team_name(bank, tid)} for tid in ids]
