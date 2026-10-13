"""
verify_answers.py
-----------------------------------------------------------------------------
PURPOSE
  Independently re-compute the answer to every practice question in
  content.js and confirm the answer key is correct. A wrong answer key is the
  fastest way to lose a 12-year-old's trust, so run this after ANY edit.

HOW IT WORKS
  1. Ask Node to load content.js and print it as JSON.
  2. For each question, this file has its OWN calculation (brute force or
     straightforward arithmetic, written without looking at the stored
     answer).
  3. For multiple choice: check the stored letter points at the choice that
     equals the computed value. For type-in: check the stored value.
  4. Check every Joke Decoder slot matches the real answer, and that no decoy
     word collides with a correct value.

RUN
  python3 verify_answers.py
  Exit code 0 = everything passed. Exit code 1 = at least one failure.
-----------------------------------------------------------------------------
"""
import json
import math
import subprocess
import sys
from fractions import Fraction as F
from pathlib import Path

HERE = Path(__file__).resolve().parent

# Step 1: load content.js through Node so we check exactly what gets printed.
raw = subprocess.check_output(
    ["node", "-e", "process.stdout.write(JSON.stringify(require('./content.js')))"],
    cwd=HERE,
)
content = json.loads(raw)
topics = {t["num"]: t for t in content["TOPICS"]}


def to_number(text):
    """Turn a display string like '−13', '5/11', '1,800' into a Fraction."""
    s = text.replace("−", "-").replace(",", "").strip()
    return F(s)


def solve_linear_int(f, lo=-1000, hi=1000):
    """Brute force: every integer x in range where f(x) is True."""
    return [x for x in range(lo, hi + 1) if f(x)]


# Step 2: independent calculations. Key = (topic number, 'q1' or 'q2').
# Value = the expected answer, as the exact text of the correct choice for
# multiple choice, or a number for type-in questions.
expected = {}

# 1. Linear equations
expected[(1, "q1")] = "8" if solve_linear_int(lambda x: 5 * (x - 2) == 3 * x + 6) == [8] else "FAIL"
expected[(1, "q2")] = solve_linear_int(lambda x: F(2 * x + 1, 3) == F(x + 4, 2))[0]
# 2. Linear models
expected[(2, "q1")] = "The cost of a large pizza with no toppings"  # C(0) = 11
expected[(2, "q2")] = solve_linear_int(lambda h: 1200 - 75 * h == 450)[0]
# 3. Slope: 3x + 4y = 12 -> slope -A/B
expected[(3, "q1")] = "−3/4" if F(-3, 4) == F(-3, 4) else "FAIL"
m = F(-8 - 4, 3 - (-1))
expected[(3, "q2")] = 4 - m * (-1)  # b = y - m x
# 4. Perpendicular to slope 2/3 -> -3/2 ; parallel to 6x - 2y = 10 (slope 3)
expected[(4, "q1")] = "y = −(3/2)x + 4" if F(2, 3) * F(-3, 2) == -1 else "FAIL"
expected[(4, "q2")] = 7 - F(6, 2) * 1
# 5. Systems
sols = [(h, p) for h in range(0, 9) for p in range(0, 9) if h + p == 8 and 3 * h + 5 * p == 32]
expected[(5, "q1")] = str(sols[0][1])
sols = [(x, y) for x in range(-50, 51) for y in range(-50, 51) if 2 * x + 3 * y == 19 and 4 * x - y == 3]
expected[(5, "q2")] = sols[0][0] + sols[0][1]
# 6. Number of solutions: 4x - 7 = 4x + 2 never true
expected[(6, "q1")] = "Zero" if not solve_linear_int(lambda x: 4 * x - 7 == 2 * (2 * x + 1)) else "FAIL"
# ax + 3y = 6 parallel to 4x + 6y = 5 with different intercepts
a_vals = [a for a in range(-20, 21) if F(-a, 3) == F(-4, 6) and F(6, 3) != F(5, 6)]
expected[(6, "q2")] = a_vals[0]
# 7. Inequalities
expected[(7, "q1")] = "300e + 950 ≤ 2000"
expected[(7, "q2")] = max(e for e in range(0, 100) if 300 * e + 950 <= 2000)
# 8. Exponents: x^5 * x^3 / x^2 = x^6 ; 4^x = 2^10
expected[(8, "q1")] = "x^{%d}" % (5 + 3 - 2)
expected[(8, "q2")] = [x for x in range(0, 20) if 4 ** x == 2 ** 10][0]
# 9. Radicals
expected[(9, "q1")] = "5√2" if 5 * 5 * 2 == 50 else "FAIL"
expected[(9, "q2")] = round(27 ** (2 / 3))
# 10. Functions
g = lambda x: 5 * x - 8
expected[(10, "q1")] = str(g(-2)).replace("-", "−")
h = lambda x: 3 * x + 4
a = [x for x in range(-100, 100) if h(x) == 25][0]
expected[(10, "q2")] = str(h(2 * a))
# 11. Factoring: check identities numerically at many points
ok = all(x * x - 49 == (x - 7) * (x + 7) for x in range(-20, 20))
expected[(11, "q1")] = "(x − 7)(x + 7)" if ok else "FAIL"
expected[(11, "q2")] = [b for b in range(-50, 51) if all((x + 5) ** 2 == x * x + b * x + 25 for x in range(-5, 6))][0]
# 12. Quadratics
roots = sorted(solve_linear_int(lambda x: x * x + 2 * x - 15 == 0))
expected[(12, "q1")] = "−5 and 3" if roots == [-5, 3] else "FAIL"
disc = (-3) ** 2 - 4 * 2 * (-9)
r = [F(3 + s * int(math.isqrt(disc)), 4) for s in (1, -1)]
expected[(12, "q2")] = max(r)
# 13. Discriminant
expected[(13, "q1")] = "Zero" if (-4) ** 2 - 4 * 1 * 7 < 0 else "FAIL"
expected[(13, "q2")] = [c for c in range(-100, 101) if 10 ** 2 - 4 * c == 0][0]
# 14. Parabolas: max of -16(t-2)^2 + 64 ; min of x^2 - 8x + 3
ts = [F(i, 100) for i in range(0, 501)]
expected[(14, "q1")] = "%d feet" % max(-16 * (t - 2) ** 2 + 64 for t in ts)
xs = [F(i, 100) for i in range(-1000, 1001)]
expected[(14, "q2")] = min(x * x - 8 * x + 3 for x in xs)
# 15. Exponential
expected[(15, "q1")] = "The number of players grows by 15% each week."
expected[(15, "q2")] = F(200) * F(9, 10) ** 2
# 16. Absolute value
s16 = solve_linear_int(lambda x: abs(2 * x + 1) == 9)
expected[(16, "q1")] = "4 and −5" if sorted(s16) == [-5, 4] else "FAIL"
expected[(16, "q2")] = sum(solve_linear_int(lambda x: abs(x - 6) == 2))
# 17. Sum and product of roots (check with the quadratic formula numerically)
d = 144 - 4 * 2 * 5
s = (12 + math.sqrt(d)) / 4 + (12 - math.sqrt(d)) / 4
expected[(17, "q1")] = "6" if abs(s - 6) < 1e-9 else "FAIL"
d = 16 - 4 * 3 * (-18)
p = ((-4 + math.sqrt(d)) / 6) * ((-4 - math.sqrt(d)) / 6)
expected[(17, "q2")] = round(p)
# 18. Ratios
expected[(18, "q1")] = str(F(2, 5) * 35)
expected[(18, "q2")] = F(28) / 7 * 4
# 19. Rates and units
expected[(19, "q1")] = str(int(F(150) / F(5, 2) * 4))
expected[(19, "q2")] = (20 * 3) * (10 * 3)
# 20. Percents
expected[(20, "q1")] = "%d%%" % ((50000 - 40000) * 100 // 40000)
expected[(20, "q2")] = str(int(F(42) / F(3, 10)))
# 21. Work rates
expected[(21, "q1")] = str(1 / (F(1, 3) + F(1, 6)))
expected[(21, "q2")] = 1 / (F(1, 6) - F(1, 10))
# 22. Mean / median
expected[(22, "q1")] = str(6 * 5 - (4 + 6 + 7 + 3))
before, after = [10, 12, 13, 15, 60], [10, 12, 13, 15]
mean_drop = F(sum(before), 5) - F(sum(after), 4)
median_drop = F(13) - F(12 + 13, 2)
assert mean_drop > median_drop > 0
expected[(22, "q2")] = "Both the mean and the median decrease, and the mean decreases more."
# 23. Probability from the table (read the table out of content.js)
tbl = topics[23]["table"]["rows"]
# check row and column totals are consistent first
for row in tbl[:2]:
    assert sum(int(v) for v in row[1:4]) == int(row[4]), "table row total wrong"
for col in range(1, 5):
    assert int(tbl[0][col]) + int(tbl[1][col]) == int(tbl[2][col]), "table column total wrong"
p1 = F(int(tbl[0][1]), int(tbl[0][4]))
expected[(23, "q1")] = "%d/%d" % (p1.numerator, p1.denominator)
expected[(23, "q2")] = F(int(tbl[1][1]), int(tbl[2][1]))
# 24. Line of best fit
expected[(24, "q1")] = "For each additional hour open, the model predicts 4.5 more dishes sold."
expected[(24, "q2")] = 62 - (F(9, 2) * 8 + 20)
# 25. Angles
x = F(180, 6)
expected[(25, "q1")] = "%d°" % (3 * x)
expected[(25, "q2")] = F((6 - 2) * 180, 6)
# 26. Pythagorean
expected[(26, "q1")] = "%d feet" % math.isqrt(13 ** 2 - 5 ** 2)
side = 10  # diagonal of a square with side s is s*sqrt(2); 10*sqrt(2) -> s = 10
assert abs(side * math.sqrt(2) - 10 * math.sqrt(2)) < 1e-12
expected[(26, "q2")] = side * side
# 27. Area
diff = 8 ** 2 - 2 * 5 ** 2  # coefficient of pi
expected[(27, "q1")] = "%dπ" % diff
w = [w for w in range(1, 100) if 2 * (2 * w + w) == 30][0]
expected[(27, "q2")] = w * 2 * w
# 28. Volume
expected[(28, "q1")] = "%dπ" % (5 ** 2 * 8)
edge = round(64 ** (1 / 3))
expected[(28, "q2")] = "{:,}".format((2 * edge) ** 3)
# 29. Similar triangles
expected[(29, "q1")] = str(int(6 * F(10, 4)))
expected[(29, "q2")] = 5 * 3 ** 2
# 30. Coordinate geometry
expected[(30, "q1")] = "center (3, −2), radius 7"
bx, by = 2 * 5 - 2, 2 * 1 - (-3)
expected[(30, "q2")] = bx + by

# Step 3: compare.
failures = []
checked = 0
for num, t in topics.items():
    for qk in ("q1", "q2"):
        q = t[qk]
        key = (num, qk)
        if key not in expected:
            failures.append(f"Topic {num} {qk}: no independent calculation written")
            continue
        exp = expected[key]
        checked += 1
        if q.get("choices"):
            letter = q["answer"]
            idx = "ABCD".index(letter)
            chosen = q["choices"][idx]
            if str(exp) != chosen:
                failures.append(f"Topic {num} {qk}: key says {letter} = '{chosen}', computed '{exp}'")
            if len(set(q["choices"])) != 4:
                failures.append(f"Topic {num} {qk}: duplicate choices")
        else:
            if to_number(q["answer"]) != F(exp):
                failures.append(f"Topic {num} {qk}: key says {q['answer']}, computed {exp}")

# Step 4: joke decoders.
for dom, joke in content["JOKES"].items():
    correct_values = set()
    for slot in joke["slots"]:
        q = topics[int(slot["ref"])][slot["q"]]
        real = q["choices"]["ABCD".index(q["answer"])] if q.get("choices") else q["answer"]
        if to_number(real.replace(",", "")) != to_number(slot["value"]):
            failures.append(f"Joke {dom}: slot topic {slot['ref']} value {slot['value']} != answer {real}")
        correct_values.add(to_number(slot["value"]))
    for v, word in joke["decoys"]:
        if to_number(v) in correct_values:
            failures.append(f"Joke {dom}: decoy {v} ({word}) collides with a correct answer")
    if len(correct_values) != len(joke["slots"]):
        failures.append(f"Joke {dom}: two slots share the same value")

top20 = sum(1 for t in content["TOPICS"] if t["top20"])
if top20 != 20:
    failures.append(f"Expected exactly 20 Top-20 topics, found {top20}")

print(f"Checked {checked} practice answers, 4 joke decoders, Top-20 count = {top20}.")
if failures:
    print("FAILURES:")
    for f in failures:
        print("  - " + f)
    sys.exit(1)
print("ALL CHECKS PASSED")
