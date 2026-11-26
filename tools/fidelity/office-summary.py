#!/usr/bin/env python3
# SPDX-License-Identifier: AGPL-3.0-only
"""Summarise a differential-batch output directory for xlsx or pptx pairs (Office desktop reference vs app PDF).
usage: office-summary.py FORMAT RESDIR OUT.md [--lowsignal-pages 30]"""
import collections, gzip, json, statistics, sys
from pathlib import Path

fmt, res, out = sys.argv[1], Path(sys.argv[2]), Path(sys.argv[3])
LOW = 30

def glyph_metrics(case):
    """Per-page character-multiset overlap and one-to-one same-character placement within 5 pt (reference glyph -> app glyph).
    Works when the unique-token metric is empty (repeated numbers in spreadsheets)."""
    try:
        g = json.loads(gzip.open(case / "glyphs.json.gz", "rt", encoding="utf-8").read())
    except OSError:
        return None, None
    def bypage(gl):
        d = collections.defaultdict(lambda: collections.defaultdict(list))
        for x in gl:
            if x["text"].strip():
                b = x["box"]; d[x["page"]][x["text"]].append(((b[0] + b[2]) / 2, (b[1] + b[3]) / 2))
        return d
    R, C = bypage(g["reference"]), bypage(g["candidate"])
    tot = same = 0; near = {5: 0, 12: 0}
    for page, chars in R.items():
        for ch, pts in chars.items():
            tot += len(pts)
            cand = list(C.get(page, {}).get(ch, []))
            same += min(len(pts), len(cand))
            for tol in (5, 12):
                pool = cand[:4000]
                for (x, y) in pts[:4000]:
                    best, bi = tol * tol, -1
                    for i, (cx, cy) in enumerate(pool):
                        d2 = (cx - x) ** 2 + (cy - y) ** 2
                        if d2 < best: best, bi = d2, i
                    if bi >= 0: near[tol] += 1; pool.pop(bi)
    return (same / tot if tot else None), ((near[5] / tot, near[12] / tot) if tot else None)

rows = []
for p in sorted(res.glob("*/result.json")):
    r = json.loads(p.read_text())
    pc = r.get("pageCounts") or {}
    t = r.get("text") or {}
    f = r.get("fonts") or {}
    ras = r.get("raster") or []
    cnt = max(pc.get("reference") or 0, pc.get("candidate") or 0)
    sim = (sum(1 - x["diffInk"] if x.get("pageBoxWithinTolerance") else 0 for x in ras) / cnt) if cnt else None
    ratio = (min(pc.get("reference") or 0, pc.get("candidate") or 0) / cnt) if cnt else None
    chov, pl = glyph_metrics(p.parent); place, place12 = pl if pl else (None, None)
    sc = r.get("diagnosticScore"); fb = False
    if sc is None and chov is not None and sim is not None and ratio is not None:
        sc = round(100 * (0.3 * ratio + 0.4 * chov + 0.3 * sim), 3); fb = True
    if sc is None and chov is None and sim is not None and ratio is not None:
        sc = round(100 * (0.3 * ratio + 0.7 * sim), 3); fb = "raster"
    rows.append(dict(id=r["id"], status=r["status"], score=sc, fallback=fb, charOverlap=chov, place5=place, place12=place12, raster=sim,
                     ref=pc.get("reference"), cand=pc.get("candidate"),
                     agree=t.get("agreement"), cov=t.get("coverage"),
                     reffonts=sorted(f.get("reference", {})), candfonts=sorted(f.get("candidate", {})),
                     firstdiv=t.get("firstPageDivergence")))
ok = [r for r in rows if r["status"] == "compared" and r["score"] is not None]
noscore = [r["id"][:10] for r in rows if r["status"] == "compared" and r["score"] is None]
def stat(xs):
    return (round(statistics.mean(xs), 2), round(statistics.median(xs), 2)) if xs else (None, None)
def block(name, rs):
    s = [r["score"] for r in rs]; m = stat(s)
    pm = sum(1 for r in rs if r["ref"] == r["cand"])
    return (f"- {name}: n={len(rs)}, score mean {m[0]} / median {m[1]}, >=99: {sum(x>=99 for x in s)}, <90: {sum(x<90 for x in s)}, "
            f"page count equal: {pm}/{len(rs)}, mean token agreement {stat([r['agree'] for r in rs if r['agree'] is not None])[0]}, "
            f"mean token coverage {stat([r['cov'] for r in rs if r['cov'] is not None])[0]}, "
            f"mean char overlap {stat([r['charOverlap'] for r in rs if r['charOverlap'] is not None])[0]}, mean glyph placement within 5pt {stat([r['place5'] for r in rs if r['place5'] is not None])[0]} / 12pt {stat([r['place12'] for r in rs if r['place12'] is not None])[0]}, "
            f"mean raster similarity {stat([r['raster'] for r in rs if r['raster'] is not None])[0]}")
L = [f"## {fmt} summary ({res})", "", f"compared {len(ok)} of {len(rows)} ({', '.join(sorted({r['status'] for r in rows if r['status']!='compared'})) or 'no failures'})", ""]
L.append(block("all compared", ok))
L.append(block("text-bearing documents only (harness or char-overlap score)", [r for r in ok if r["fallback"] != "raster"]))
L.append(f"- score marked * = harness score unavailable (no unique reference token, e.g. repeated numbers); same formula with per-page character overlap instead of unique-token agreement: {sum(r['fallback'] is True for r in ok)} docs. Score marked \u2020 = the reference PDF has no text at all (shapes/pictures only), score = 100*(0.3*page ratio + 0.7*raster similarity): {sum(r['fallback']=='raster' for r in ok)} docs")
if fmt == "xlsx":
    sig = [r for r in ok if (r["ref"] or 0) <= LOW]
    low = [r for r in ok if (r["ref"] or 0) > LOW]
    L.append(block(f"high-signal (Excel pages <= {LOW})", sig)); L.append(f"- low-signal (Excel prints > {LOW} pages): {len(low)} documents: " + ", ".join(r["id"][:10] for r in low))
else:
    L.append(f"- slide count exact match: {sum(r['ref']==r['cand'] for r in ok)}/{len(ok)}; mismatches: " +
             ", ".join(f"{r['id'][:10]}({r['ref']} vs {r['cand']})" for r in ok if r["ref"] != r["cand"]))
if noscore: L.append(f"- no score at all (no reference text, blank or pictures only): {len(noscore)}: " + ", ".join(noscore))
L += ["", "### worst 10 (by diagnostic score)", "", "| id | score | pages ref/app | token agree | coverage | char overlap | place<5pt/12pt | reference fonts | app fonts |", "|---|---|---|---|---|---|---|---|---|"]
for r in sorted(ok, key=lambda r: r["score"])[:10]:
    clean = lambda fs: ", ".join(x.split("+", 1)[-1] for x in fs)[:70]
    L.append(f"| {r['id'][:12]} | {r['score']}{'*' if r['fallback'] is True else ('\u2020' if r['fallback'] else '')} | {r['ref']}/{r['cand']} | {r['agree'] if r['agree'] is None else round(r['agree'],2)} | {r['cov'] if r['cov'] is None else round(r['cov'],2)} | {round(r['charOverlap'],2) if r['charOverlap'] is not None else None} | {round(r['place5'],2) if r['place5'] is not None else None}/{round(r['place12'],2) if r['place12'] is not None else None} | {clean(r['reffonts'])} | {clean(r['candfonts'])} |")
out.write_text("\n".join(L) + "\n")
(out.with_suffix(".json")).write_text(json.dumps(rows, indent=1))
print("\n".join(L[:9]))
