# PSAT 8/9 Math Guide

A 23-page printable study guide for middle school math club students preparing for the PSAT 8/9.

**Open this:** `PSAT_8-9_Math_Guide.docx`

## What is inside

| Pages | Section |
|-------|---------|
| 1-2   | Formula sheet by test domain, with the 20 most frequent topics tagged TOP 20; how the guide works; Top 20 checklist |
| 3-18  | 30 topics. Each one: the idea, a worked example, Question 1 (Normal), Question 2 (Boss level), and a "Watch out" trap. Each domain ends with a Spot the Mistake and a Joke Decoder |
| 19-23 | Answer key with step-by-step solutions (two columns) |

Logarithms and trigonometry are left out on purpose: the PSAT 8/9 does not test them.

## Files

| File | Role |
|------|------|
| `content.js` | ALL math content (topics, questions, answers, solutions, jokes). Edit this to change the guide. |
| `build_guide.js` | Layout only. Reads `content.js` and writes the .docx. |
| `verify_answers.py` | Recomputes every answer independently and checks the answer key and joke decoders. |
| `PSAT_8-9_Math_Guide.docx` | The generated guide. |

## Rebuilding after an edit

```bash
cd psat-8-9-guide
python3 verify_answers.py     # must print ALL CHECKS PASSED
node build_guide.js           # writes PSAT_8-9_Math_Guide.docx
```

`build_guide.js` needs the `docx` npm package (`npm install docx` if it is missing).
