/*
 * content.js
 * ----------------------------------------------------------------------------
 * PURPOSE
 *   This file holds ALL of the math content for the PSAT 8/9 Math Guide:
 *   the formula sheet, the 30 topic blocks, the "Spot the Mistake" problems,
 *   and the "Joke Decoder" puzzles. It contains NO layout code.
 *
 *   build_guide.js reads this file and turns it into the .docx document.
 *   verify_answers.py re-computes every answer independently and checks it
 *   against the answers stored here (run it after ANY content change).
 *
 * HOW THE PIECES CONNECT
 *   content.js  --(require)-->  build_guide.js  -->  PSAT_8-9_Math_Guide.docx
 *   content.js  --(node dumps JSON)-->  verify_answers.py  -->  PASS / FAIL
 *
 * MINI MARKUP USED INSIDE STRINGS (build_guide.js understands these)
 *   ^{...}   superscript   example: "x^{2}"  renders as x with a raised 2
 *   _{...}   subscript     example: "y_{1}"  renders as y with a lowered 1
 *   **...**  bold          example: "**Step 1:**"
 *   Everything else is plain Unicode text (minus sign, sqrt sign, pi, etc).
 *
 * DATA SHAPE OF ONE TOPIC
 *   {
 *     num:     topic number shown in the guide (1..30)
 *     domain:  'A' Algebra | 'M' Advanced Math | 'D' Data | 'G' Geometry
 *     title:   topic name
 *     top20:   true if this topic is in the "Top 20 - shows up almost every test"
 *     bonus:   true if this is an optional / stretch topic
 *     idea:    one or two short sentences: the concept in plain words
 *     formulas: optional list of key formulas for this topic
 *     table:   optional data table { header: [...], rows: [[...], ...] }
 *     example: { problem, steps[] }  worked example shown in the topic block
 *     q1, q2:  practice questions:
 *              { stem, choices (4 strings, multiple choice) OR none (type-in),
 *                answer (letter for multiple choice, value string for type-in),
 *                solution: [step strings shown in the Answer Key at the back] }
 *              q1 is "Normal level", q2 is "Boss level" (a bit harder).
 *     watch:   the most common trap for this topic
 *     desmos:  optional one-line calculator shortcut
 *   }
 *
 * NAMES: all names in word problems are made up (by request: safer to share).
 * ----------------------------------------------------------------------------
 */

// ---------------------------------------------------------------------------
// Domain metadata: labels and colors used by build_guide.js.
// Colors are muted so the guide still reads well when printed in black/white.
//   dark  = header text / rule color
//   tint  = light background shading for the topic header bar
// ---------------------------------------------------------------------------
const DOMAINS = {
  A: { name: 'Algebra', dark: '1F3A68', tint: 'DCE6F2', share: 'about 35% of the test' },
  M: { name: 'Advanced Math', dark: '0F5953', tint: 'D6ECE9', share: 'about 32% of the test' },
  D: { name: 'Problem-Solving & Data', dark: '7A2E2E', tint: 'F3E1DD', share: 'about 20% of the test' },
  G: { name: 'Geometry', dark: '4A5A1E', tint: 'E6EDD6', share: 'about 13% of the test' },
};

// ---------------------------------------------------------------------------
// FORMULA SHEET (pages 1-2).
// Each line: { t: text with markup, n: topic number it links to, top: Top 20? }
// "given: true" lines are on the official reference sheet on test day.
// ---------------------------------------------------------------------------
const FORMULA_SHEET = {
  A: [
    { t: 'Solve an equation: undo + and − first, then × and ÷. Do the same to both sides.', n: 1, top: true },
    { t: 'Linear model: y = mx + b.  m = rate (change per 1 unit),  b = starting value (when x = 0)', n: 2, top: true },
    { t: 'Slope: m = (y_{2} − y_{1}) / (x_{2} − x_{1})     Point-slope: y − y_{1} = m(x − x_{1})', n: 3, top: true },
    { t: 'Standard form: Ax + By = C has slope −A/B', n: 3, top: true },
    { t: 'Parallel lines: same slope.  Perpendicular: slopes multiply to −1  (2 and −1/2)', n: 4 },
    { t: 'Systems: the (x, y) that works in BOTH equations. Substitute, eliminate, or graph.', n: 5, top: true },
    { t: 'Different slopes: 1 solution.  Same slope, different intercept: NO solution.  Same line: infinitely many.', n: 6, top: true },
    { t: 'Inequalities: multiply or divide by a negative, FLIP the sign.  "at least" ≥,  "at most" ≤', n: 7, top: true },
  ],
  M: [
    { t: 'a^{m} · a^{n} = a^{m+n}     a^{m} / a^{n} = a^{m−n}     (a^{m})^{n} = a^{mn}', n: 8, top: true },
    { t: 'a^{0} = 1     a^{−n} = 1 / a^{n}     (ab)^{n} = a^{n}b^{n}', n: 8, top: true },
    { t: 'a^{1/n} = ⁿ√a     a^{m/n} = (ⁿ√a)^{m}     √(ab) = √a · √b', n: 9 },
    { t: 'Function notation: f(3) means "put 3 in for x."', n: 10, top: true },
    { t: '(a + b)^{2} = a^{2} + 2ab + b^{2}     (a − b)^{2} = a^{2} − 2ab + b^{2}     a^{2} − b^{2} = (a + b)(a − b)', n: 11, top: true },
    { t: 'Quadratic formula: x = [−b ± √(b^{2} − 4ac)] / 2a   (set the equation = 0 first)', n: 12, top: true },
    { t: 'Discriminant b^{2} − 4ac:  positive → 2 real solutions,  zero → 1,  negative → 0', n: 13 },
    { t: 'Vertex form: y = a(x − h)^{2} + k, vertex (h, k).  In standard form, vertex x = −b / 2a', n: 14, top: true },
    { t: 'Exponential: y = a(1 + r)^{t} growth,  y = a(1 − r)^{t} decay   (r as a decimal)', n: 15, top: true },
    { t: '|x − a| = b  →  x − a = b  OR  x − a = −b', n: 16 },
    { t: 'Bonus: for ax^{2} + bx + c = 0, sum of roots = −b/a, product of roots = c/a', n: 17 },
  ],
  D: [
    { t: 'Proportion: a/b = c/d  →  ad = bc   (check: part to part, or part to whole?)', n: 18, top: true },
    { t: 'distance = rate × time.  Convert units by multiplying fractions that cancel.', n: 19, top: true },
    { t: 'Part = (percent / 100) × whole     Percent change = (new − old) / old × 100%', n: 20, top: true },
    { t: 'Increase by p%: multiply by (1 + p/100).  Decrease by p%: multiply by (1 − p/100)', n: 20, top: true },
    { t: 'Working together: rate = 1 / time.  ADD the rates, never the times.', n: 21 },
    { t: 'Mean = sum / count.  Sum = mean × count.  Median = middle (in order).  Range = max − min', n: 22, top: true },
    { t: 'P(event) = favorable / total     P(not A) = 1 − P(A)     P(A or B) = P(A) + P(B) − P(A and B)', n: 23, top: true },
    { t: '"Given" in a two-way table: only that row or column is the total.', n: 23, top: true },
    { t: 'Line of best fit y = mx + b: slope = predicted change per 1 unit of x', n: 24 },
  ],
  G: [
    { t: 'Straight line = 180°.  Triangle = 180°.  Around a point = 360°.  Vertical angles are equal.', n: 25, top: true },
    { t: 'Polygon with n sides: interior angles add to (n − 2) × 180°', n: 25, top: true },
    { t: 'Pythagorean: a^{2} + b^{2} = c^{2}.  Triples: 3-4-5, 5-12-13, 8-15-17, 6-8-10', n: 26, top: true },
    { t: 'Special right triangles: 45-45-90 → x, x, x√2     30-60-90 → x, x√3, 2x', n: 26, top: true },
    { t: 'Rectangle A = lw   Triangle A = ½bh   Circle A = πr^{2}, C = 2πr', n: 27, top: true },
    { t: 'Box V = lwh   Cylinder V = πr^{2}h   Sphere V = (4/3)πr^{3}   Cone V = (1/3)πr^{2}h', n: 28 },
    { t: 'Similar shapes: lengths × k, areas × k^{2}, volumes × k^{3}', n: 29 },
    { t: 'Distance √[(x_{2} − x_{1})^{2} + (y_{2} − y_{1})^{2}]   Midpoint ((x_{1} + x_{2})/2, (y_{1} + y_{2})/2)', n: 30 },
    { t: 'Circle: (x − h)^{2} + (y − k)^{2} = r^{2}, center (h, k), radius r', n: 30 },
  ],
};

// Things the test DOES NOT need (shown on the formula sheet so nobody wastes time).
const NOT_ON_TEST = 'Not on the PSAT 8/9: logarithms and trigonometry (sin, cos, tan). Skip them for now.';

// ---------------------------------------------------------------------------
// THE 30 TOPICS
// ---------------------------------------------------------------------------
const TOPICS = [
  // ======================== ALGEBRA ========================
  {
    num: 1, domain: 'A', title: 'Solving Linear Equations', top20: true,
    idea: 'Keep the equation balanced: whatever you do to one side, do to the other. Undo adding and subtracting first, then multiplying and dividing.',
    example: {
      problem: 'Solve 3x + 7 = 31.',
      steps: ['Subtract 7 from both sides: 3x = 24.', 'Divide both sides by 3: x = 8.', 'Check: 3(8) + 7 = 24 + 7 = 31. It works.'],
    },
    q1: {
      stem: 'What value of x satisfies 5(x − 2) = 3x + 6 ?',
      choices: ['2', '4', '8', '16'], answer: 'C',
      solution: [
        'Distribute the 5 to BOTH terms inside: 5x − 10 = 3x + 6.',
        'Subtract 3x from both sides: 2x − 10 = 6.',
        'Add 10 to both sides: 2x = 16.',
        'Divide by 2: x = 8.',
        'Trap: forgetting to multiply the −2 by 5 gives x = 4 (choice B).',
      ],
    },
    q2: {
      stem: 'If (2x + 1) / 3 = (x + 4) / 2, what is the value of x?',
      answer: '10',
      solution: [
        'Cross-multiply (multiply each top by the other bottom): 2(2x + 1) = 3(x + 4).',
        'Distribute: 4x + 2 = 3x + 12.',
        'Subtract 3x: x + 2 = 12.',
        'Subtract 2: x = 10.',
      ],
    },
    watch: 'A number in front of parentheses multiplies EVERY term inside, not just the first one.',
  },
  {
    num: 2, domain: 'A', title: 'Linear Models: What the Numbers Mean', top20: true,
    idea: 'In y = mx + b, the number multiplied by x (m) is the rate: how much y changes each time x goes up by 1. The number by itself (b) is the starting amount, when x = 0.',
    example: {
      problem: 'Ava has 150 Robux and earns 40 Robux each week doing chores. Write an equation for her Robux R after w weeks, then find R after 6 weeks.',
      steps: ['Start = 150, rate = 40 per week, so R = 40w + 150.', 'After 6 weeks: R = 40(6) + 150 = 240 + 150 = 390 Robux.'],
    },
    q1: {
      stem: 'A pizza shop charges C = 2.5t + 11 dollars for a large pizza with t toppings. What does the number 11 represent?',
      choices: ['The cost of each topping', 'The cost of a large pizza with no toppings', 'The total cost of a pizza with 11 toppings', 'The number of toppings on the pizza'],
      answer: 'B',
      solution: [
        'The 11 is the number by itself, so it is the starting value.',
        'It is the cost when t = 0: C = 2.5(0) + 11 = 11.',
        'So $11 is the cost of a large pizza with zero toppings.',
        'The 2.5 (attached to t) is the cost of each topping, which is choice A. That is the trap.',
      ],
    },
    q2: {
      stem: 'A Minecraft server has 1,200 players online at noon and loses 75 players every hour. After how many hours will it have 450 players online?',
      answer: '10',
      solution: [
        'Start at 1,200 and lose 75 each hour: P = 1200 − 75h.',
        'Set P = 450: 1200 − 75h = 450.',
        'Subtract 1200 from both sides: −75h = −750.',
        'Divide by −75: h = 10 hours.',
      ],
    },
    watch: 'The rate is the number stuck to the variable. The number standing alone is the start.',
  },
  {
    num: 3, domain: 'A', title: 'Slope and the Forms of a Line', top20: true,
    formulas: ['m = (y_{2} − y_{1}) / (x_{2} − x_{1})', 'y = mx + b     y − y_{1} = m(x − x_{1})     Ax + By = C'],
    idea: 'Slope is "rise over run": how much y changes for each step of 1 in x.',
    example: {
      problem: 'Find the equation of the line through (2, 5) and (6, 13).',
      steps: ['Slope: m = (13 − 5) / (6 − 2) = 8 / 4 = 2.', 'Point-slope with (2, 5): y − 5 = 2(x − 2).', 'Simplify: y = 2x − 4 + 5, so y = 2x + 1.'],
    },
    q1: {
      stem: 'What is the slope of the line 3x + 4y = 12 ?',
      choices: ['−4/3', '−3/4', '3/4', '3'], answer: 'B',
      solution: [
        'Get y alone. Subtract 3x: 4y = −3x + 12.',
        'Divide everything by 4: y = −(3/4)x + 3.',
        'The number in front of x is the slope: −3/4.',
        'Shortcut: in Ax + By = C, slope = −A/B = −3/4.',
      ],
    },
    q2: {
      stem: 'A line passes through (−1, 4) and (3, −8). What is the y-intercept of the line?',
      answer: '1',
      solution: [
        'Slope: m = (−8 − 4) / (3 − (−1)) = −12 / 4 = −3.',
        'So y = −3x + b. Plug in the point (−1, 4): 4 = −3(−1) + b.',
        'That is 4 = 3 + b, so b = 1.',
        'Check with (3, −8): −3(3) + 1 = −8. It works.',
      ],
    },
    watch: 'Subtract in the SAME order on top and bottom. And 3 − (−1) is 3 + 1 = 4.',
  },
  {
    num: 4, domain: 'A', title: 'Parallel and Perpendicular Lines', top20: false,
    idea: 'Parallel lines have the same slope. Perpendicular lines have slopes that are "flipped and negated": 2 goes with −1/2.',
    example: {
      problem: 'Find the line through (0, 3) that is parallel to y = 4x − 7.',
      steps: ['Parallel means same slope: m = 4.', 'The point (0, 3) is the y-intercept, so b = 3.', 'Answer: y = 4x + 3.'],
    },
    q1: {
      stem: 'Which line is perpendicular to y = (2/3)x + 5 ?',
      choices: ['y = (2/3)x − 1', 'y = −(2/3)x + 5', 'y = (3/2)x + 2', 'y = −(3/2)x + 4'], answer: 'D',
      solution: [
        'The slope of the given line is 2/3.',
        'Flip it: 3/2. Change the sign: −3/2.',
        'Check: (2/3) × (−3/2) = −1. Perpendicular slopes always multiply to −1.',
        'Only choice D has slope −3/2.',
      ],
    },
    q2: {
      stem: 'Line k is parallel to 6x − 2y = 10 and passes through (1, 7). What is the y-intercept of line k?',
      answer: '4',
      solution: [
        'Find the slope of 6x − 2y = 10. Subtract 6x: −2y = −6x + 10.',
        'Divide by −2: y = 3x − 5. The slope is 3.',
        'Line k has slope 3 too: y = 3x + b.',
        'Plug in (1, 7): 7 = 3(1) + b, so b = 4.',
      ],
    },
    watch: 'Perpendicular needs BOTH moves: flip the fraction AND change the sign.',
  },
  {
    num: 5, domain: 'A', title: 'Systems of Two Equations', top20: true,
    idea: 'A system asks for the one (x, y) pair that works in both equations. Use substitution, elimination (add the equations), or graph both lines.',
    desmos: 'Type both equations into Desmos and click where the lines cross.',
    example: {
      problem: 'Solve: x + y = 10 and x − y = 4.',
      steps: ['Add the two equations (the y terms cancel): 2x = 14, so x = 7.', 'Put x = 7 into x + y = 10: y = 3.', 'Solution: (7, 3).'],
    },
    q1: {
      stem: 'At the math club fundraiser, hot dogs cost $3 and pasta bowls cost $5. Kai bought 8 items for $32. How many pasta bowls did Kai buy?',
      choices: ['2', '3', '4', '6'], answer: 'C',
      solution: [
        'Let h = hot dogs and p = pasta bowls.',
        'Number of items: h + p = 8. Money: 3h + 5p = 32.',
        'From the first equation, h = 8 − p. Substitute: 3(8 − p) + 5p = 32.',
        'Distribute: 24 − 3p + 5p = 32, so 24 + 2p = 32.',
        'Subtract 24: 2p = 8, so p = 4 pasta bowls (and h = 4 hot dogs).',
        'Check: 3(4) + 5(4) = 12 + 20 = 32. It works.',
      ],
    },
    q2: {
      stem: 'If 2x + 3y = 19 and 4x − y = 3, what is the value of x + y ?',
      answer: '7',
      solution: [
        'Solve the second equation for y: y = 4x − 3.',
        'Substitute into the first: 2x + 3(4x − 3) = 19.',
        'Distribute: 2x + 12x − 9 = 19, so 14x − 9 = 19.',
        'Add 9: 14x = 28, so x = 2.',
        'Then y = 4(2) − 3 = 5.',
        'The question asks for x + y = 2 + 5 = 7.',
      ],
    },
    watch: 'Read the last line of the question. If it asks for x + y, do not stop at x.',
  },
  {
    num: 6, domain: 'A', title: 'One, None, or Infinitely Many Solutions', top20: true,
    idea: 'If x cancels out and you get something TRUE (like 6 = 6), there are infinitely many solutions. Something FALSE (like −7 = 2) means no solution. For two lines: same slope and different intercept means no solution (parallel lines never meet).',
    example: {
      problem: 'How many solutions does 3(x + 2) = 3x + 6 have?',
      steps: ['Distribute: 3x + 6 = 3x + 6.', 'Both sides are exactly the same, so every x works.', 'Infinitely many solutions.'],
    },
    q1: {
      stem: 'How many solutions does 4x − 7 = 2(2x + 1) have?',
      choices: ['Zero', 'Exactly one', 'Exactly two', 'Infinitely many'], answer: 'A',
      solution: [
        'Distribute on the right: 4x − 7 = 4x + 2.',
        'Subtract 4x from both sides: −7 = 2.',
        'That is false, and x is gone, so no value of x can make it true.',
        'Zero solutions.',
      ],
    },
    q2: {
      stem: 'In the system ax + 3y = 6 and 4x + 6y = 5, a is a constant. The system has no solution. What is the value of a?',
      answer: '2',
      solution: [
        'No solution means the lines are parallel: same slope, different intercept.',
        'Slope of 4x + 6y = 5 is −A/B = −4/6 = −2/3.',
        'Slope of ax + 3y = 6 is −a/3.',
        'Set them equal: −a/3 = −2/3, so a = 2.',
        'Check the intercepts are different: 2x + 3y = 6 gives y-intercept 2; 4x + 6y = 5 gives 5/6. Different, so they never meet.',
      ],
    },
    watch: 'Same slope alone is not enough for infinitely many. The intercepts must ALSO match.',
  },
  {
    num: 7, domain: 'A', title: 'Linear Inequalities', top20: true,
    idea: 'Solve like an equation, but if you multiply or divide by a negative number, flip the inequality sign. Word clues: "at least" is ≥, "at most" or "no more than" is ≤.',
    example: {
      problem: 'Solve −2x + 5 > 13.',
      steps: ['Subtract 5: −2x > 8.', 'Divide by −2 and FLIP the sign: x < −4.'],
    },
    q1: {
      stem: 'Zoe has 2,000 V-Bucks. She buys a battle pass for 950 V-Bucks and wants to buy emotes that cost 300 V-Bucks each. Which inequality shows the number of emotes, e, she can buy?',
      choices: ['950e + 300 ≤ 2000', '300e + 950 ≤ 2000', '300e + 950 ≥ 2000', '300e ≤ 2000 + 950'], answer: 'B',
      solution: [
        'Each emote costs 300, so e emotes cost 300e.',
        'Add the battle pass: total spent = 300e + 950.',
        'She cannot spend more than she has, so total ≤ 2000.',
        'Answer: 300e + 950 ≤ 2000.',
      ],
    },
    q2: {
      stem: 'Using the situation in Question 1, what is the greatest number of emotes Zoe can buy?',
      answer: '3',
      solution: [
        'Start with 300e + 950 ≤ 2000.',
        'Subtract 950: 300e ≤ 1050.',
        'Divide by 300: e ≤ 3.5.',
        'You cannot buy half an emote, and 4 is too many, so the greatest number is 3.',
        'Check: 3 emotes cost 900, plus 950 is 1,850, which is under 2,000.',
      ],
    },
    watch: 'Dividing by a negative flips the sign. In word problems, round to a whole number that still fits.',
  },

  // ======================== ADVANCED MATH ========================
  {
    num: 8, domain: 'M', title: 'Exponent Rules', top20: true,
    formulas: ['a^{m} · a^{n} = a^{m+n}     a^{m} / a^{n} = a^{m−n}     (a^{m})^{n} = a^{mn}     a^{0} = 1     a^{−n} = 1/a^{n}'],
    idea: 'Multiplying the same base: ADD exponents. Dividing: SUBTRACT. Power of a power: MULTIPLY.',
    example: {
      problem: 'Simplify (2x^{3})^{2} · x^{4}.',
      steps: ['Square everything inside: 2^{2} · x^{3×2} = 4x^{6}.', 'Multiply by x^{4}: add exponents, 4x^{6+4} = 4x^{10}.'],
    },
    q1: {
      stem: 'Which expression is equivalent to (x^{5} · x^{3}) / x^{2} ?',
      choices: ['x^{4}', 'x^{6}', 'x^{13}', 'x^{16}'], answer: 'B',
      solution: [
        'Top: x^{5} · x^{3} = x^{5+3} = x^{8}.',
        'Divide by x^{2}: subtract exponents, x^{8−2} = x^{6}.',
      ],
    },
    q2: {
      stem: 'If 4^{x} = 2^{10}, what is the value of x?',
      answer: '5',
      solution: [
        'Make the bases match. 4 is 2^{2}, so 4^{x} = (2^{2})^{x} = 2^{2x}.',
        'Now 2^{2x} = 2^{10}. Same base, so the exponents are equal: 2x = 10.',
        'x = 5.',
        'Check: 4^{5} = 1024 and 2^{10} = 1024.',
      ],
    },
    watch: 'Never multiply the bases: 2^{3} · 2^{4} = 2^{7}, not 4^{7}.',
  },
  {
    num: 9, domain: 'M', title: 'Radicals and Fractional Exponents', top20: false,
    formulas: ['a^{1/n} = ⁿ√a     a^{m/n} = (ⁿ√a)^{m}     √(ab) = √a · √b'],
    idea: 'In a fractional exponent, the bottom number is the root and the top number is the power. To simplify a square root, pull out the biggest perfect square.',
    example: {
      problem: 'Find 8^{2/3}, and simplify √72.',
      steps: ['8^{2/3}: cube root of 8 is 2, then square it: 2^{2} = 4.', '√72 = √(36 · 2) = √36 · √2 = 6√2.'],
    },
    q1: {
      stem: 'Which of the following is equal to √50 ?',
      choices: ['2√5', '5√2', '10√5', '25√2'], answer: 'B',
      solution: [
        'Find the biggest perfect square that divides 50: 25.',
        '√50 = √(25 · 2) = √25 · √2 = 5√2.',
      ],
    },
    q2: {
      stem: 'What is the value of 27^{2/3} ?',
      answer: '9',
      solution: [
        'Bottom of the exponent (3) is the root: the cube root of 27 is 3, because 3 × 3 × 3 = 27.',
        'Top of the exponent (2) is the power: 3^{2} = 9.',
      ],
    },
    watch: 'x^{1/2} is the square root of x, not x divided by 2.',
  },
  {
    num: 10, domain: 'M', title: 'Function Notation', top20: true,
    idea: 'Think of f(x) as a machine. f(4) means "put 4 in for x." If the problem says f(x) = 7, you are finding the INPUT that gives 7.',
    example: {
      problem: 'If f(x) = 2x^{2} − 3, find f(4).',
      steps: ['Replace every x with 4: f(4) = 2(4)^{2} − 3.', 'Exponent first: 2(16) − 3 = 32 − 3 = 29.'],
    },
    q1: {
      stem: 'If g(x) = 5x − 8, what is the value of g(−2) ?',
      choices: ['−18', '−2', '2', '18'], answer: 'A',
      solution: [
        'Replace x with −2: g(−2) = 5(−2) − 8.',
        '5 times −2 is −10, so −10 − 8 = −18.',
      ],
    },
    q2: {
      stem: 'The function h is defined by h(x) = 3x + 4. If h(a) = 25, what is the value of h(2a) ?',
      choices: ['14', '29', '46', '50'], answer: 'C',
      solution: [
        'First find a. h(a) = 25 means 3a + 4 = 25.',
        'Subtract 4: 3a = 21, so a = 7.',
        'Then 2a = 14. Find h(14) = 3(14) + 4 = 42 + 4 = 46.',
        'Trap: doubling 25 to get 50 (choice D). h(2a) is NOT 2 × h(a).',
      ],
    },
    watch: 'f(2a) is not the same as 2 × f(a). Find a first, then plug in.',
  },
  {
    num: 11, domain: 'M', title: 'Factoring and Multiplying Out', top20: true,
    formulas: ['(a + b)^{2} = a^{2} + 2ab + b^{2}     a^{2} − b^{2} = (a + b)(a − b)'],
    idea: 'To factor x^{2} + bx + c, find two numbers that MULTIPLY to c and ADD to b.',
    example: {
      problem: 'Factor x^{2} + 7x + 12.',
      steps: ['Two numbers that multiply to 12 and add to 7: 3 and 4.', 'So x^{2} + 7x + 12 = (x + 3)(x + 4).', 'Check by multiplying back: x^{2} + 4x + 3x + 12. It matches.'],
    },
    q1: {
      stem: 'Which expression is equivalent to x^{2} − 49 ?',
      choices: ['(x − 7)^{2}', '(x + 7)^{2}', '(x − 7)(x + 7)', '(x − 49)(x + 1)'], answer: 'C',
      solution: [
        'x^{2} − 49 is a "difference of squares": x^{2} minus 7^{2}.',
        'The pattern a^{2} − b^{2} = (a + b)(a − b) gives (x + 7)(x − 7).',
        'Check: (x − 7)(x + 7) = x^{2} + 7x − 7x − 49 = x^{2} − 49.',
      ],
    },
    q2: {
      stem: 'If (x + 5)^{2} = x^{2} + bx + 25 for all values of x, what is the value of b?',
      answer: '10',
      solution: [
        'Write (x + 5)^{2} as (x + 5)(x + 5).',
        'Multiply each term by each term: (x)(x) + (x)(5) + (5)(x) + (5)(5).',
        'That is x^{2} + 5x + 5x + 25 = x^{2} + 10x + 25.',
        'So b = 10.',
      ],
    },
    watch: '(x + 5)^{2} is NOT x^{2} + 25. Do not forget the middle term.',
  },
  {
    num: 12, domain: 'M', title: 'Solving Quadratic Equations', top20: true,
    formulas: ['x = [−b ± √(b^{2} − 4ac)] / 2a'],
    idea: 'Move everything to one side so it equals 0. Then factor (if A × B = 0, then A = 0 or B = 0) or use the quadratic formula.',
    desmos: 'Graph y = (the expression). The x-intercepts are the solutions.',
    example: {
      problem: 'Solve x^{2} − 5x + 6 = 0.',
      steps: ['Two numbers that multiply to 6 and add to −5: −2 and −3.', '(x − 2)(x − 3) = 0.', 'So x − 2 = 0 or x − 3 = 0, which gives x = 2 or x = 3.'],
    },
    q1: {
      stem: 'What are the solutions to x^{2} + 2x − 15 = 0 ?',
      choices: ['−5 and 3', '5 and −3', '−5 and −3', '5 and 3'], answer: 'A',
      solution: [
        'Two numbers that multiply to −15 and add to +2: 5 and −3.',
        'Factor: (x + 5)(x − 3) = 0.',
        'x + 5 = 0 gives x = −5. x − 3 = 0 gives x = 3.',
        'Trap: the solutions have the OPPOSITE signs of the numbers in the factors.',
      ],
    },
    q2: {
      stem: 'What is the positive solution to 2x^{2} − 3x − 9 = 0 ?',
      answer: '3',
      solution: [
        'Use the quadratic formula with a = 2, b = −3, c = −9.',
        'b^{2} − 4ac = (−3)^{2} − 4(2)(−9) = 9 + 72 = 81, and √81 = 9.',
        'x = [3 ± 9] / 4.',
        'x = 12/4 = 3 or x = −6/4 = −1.5.',
        'The positive solution is 3.',
      ],
    },
    watch: 'Set the equation equal to 0 before factoring. Never divide both sides by x; you lose the answer x = 0.',
  },
  {
    num: 13, domain: 'M', title: 'The Discriminant: How Many Solutions?', top20: false,
    formulas: ['b^{2} − 4ac > 0: two real solutions     = 0: one     < 0: none'],
    idea: 'The part under the square root in the quadratic formula tells you how many real solutions there are, without solving.',
    example: {
      problem: 'How many real solutions does x^{2} + 6x + 9 = 0 have?',
      steps: ['a = 1, b = 6, c = 9. b^{2} − 4ac = 36 − 36 = 0.', 'Zero means exactly one solution (it is x = −3).'],
    },
    q1: {
      stem: 'How many real solutions does x^{2} − 4x + 7 = 0 have?',
      choices: ['Zero', 'Exactly one', 'Exactly two', 'Infinitely many'], answer: 'A',
      solution: [
        'a = 1, b = −4, c = 7.',
        'b^{2} − 4ac = (−4)^{2} − 4(1)(7) = 16 − 28 = −12.',
        'Negative, so there are zero real solutions.',
      ],
    },
    q2: {
      stem: 'The equation x^{2} + 10x + c = 0 has exactly one real solution. What is the value of c?',
      answer: '25',
      solution: [
        'Exactly one solution means b^{2} − 4ac = 0.',
        'a = 1, b = 10: 10^{2} − 4(1)(c) = 0, so 100 − 4c = 0.',
        'Add 4c: 100 = 4c, so c = 25.',
      ],
    },
    watch: 'b^{2} is never negative: if b = −4, then b^{2} = 16.',
  },
  {
    num: 14, domain: 'M', title: 'Parabolas and Vertex Form', top20: true,
    formulas: ['y = a(x − h)^{2} + k  →  vertex (h, k)     Standard form: vertex x = −b / 2a'],
    idea: 'The vertex is the turning point. If a is positive the parabola opens up (vertex is the minimum). If a is negative it opens down (vertex is the maximum).',
    example: {
      problem: 'What is the vertex of y = 2(x − 3)^{2} + 5, and is it a max or min?',
      steps: ['Match to a(x − h)^{2} + k: h = 3, k = 5. Vertex (3, 5).', 'a = 2 is positive, so it opens up and 5 is the minimum value.'],
    },
    q1: {
      stem: 'A Fortnite launch pad sends a player into the air. The player\'s height, in feet, t seconds after launch is h(t) = −16(t − 2)^{2} + 64. What is the maximum height?',
      choices: ['2 feet', '16 feet', '32 feet', '64 feet'], answer: 'D',
      solution: [
        'This is vertex form a(t − h)^{2} + k with a = −16, h = 2, k = 64.',
        'a is negative, so the parabola opens down and the vertex is the highest point.',
        'The maximum height is k = 64 feet (reached at t = 2 seconds).',
        'Trap: 2 is the TIME of the max, not the height.',
      ],
    },
    q2: {
      stem: 'What is the minimum value of y = x^{2} − 8x + 3 ?',
      answer: '−13',
      solution: [
        'The vertex x is −b / 2a = −(−8) / 2(1) = 8/2 = 4.',
        'The minimum value is the y at x = 4: y = 4^{2} − 8(4) + 3.',
        '16 − 32 + 3 = −13.',
        'Desmos check: graph y = x^{2} − 8x + 3 and click the lowest point: (4, −13).',
      ],
    },
    watch: 'In y = (x − 3)^{2}, the vertex is at x = +3. The sign inside flips.',
  },
  {
    num: 15, domain: 'M', title: 'Exponential Growth and Decay', top20: true,
    formulas: ['Growth: y = a(1 + r)^{t}     Decay: y = a(1 − r)^{t}     (a = start, r = rate as a decimal)'],
    idea: 'Linear adds the SAME AMOUNT each time. Exponential multiplies by the SAME NUMBER each time.',
    example: {
      problem: 'A kids\' music video has 5,000 views and its views grow 20% per day. How many views after 2 days?',
      steps: ['20% growth means multiply by 1.20 each day: V = 5000(1.2)^{d}.', 'After 2 days: 5000 × 1.2 × 1.2 = 5000 × 1.44 = 7,200 views.'],
    },
    q1: {
      stem: 'The number of players in a new Roblox game is modeled by P = 800(1.15)^{w}, where w is the number of weeks after launch. What does 1.15 tell you?',
      choices: ['The game gains 15 players each week.', 'The number of players grows by 15% each week.', 'The number of players grows by 115% each week.', 'The game started with 115 players.'],
      answer: 'B',
      solution: [
        'In a(1 + r)^{t}, the number being raised to the power is 1 + r.',
        '1.15 = 1 + 0.15, so r = 0.15 = 15%.',
        'Each week the players are multiplied by 1.15, which means a 15% increase.',
        'The 800 is the starting number of players.',
      ],
    },
    q2: {
      stem: 'A Minecraft potion has a strength of 200 and loses 10% of its strength every minute. What is its strength after 2 minutes?',
      answer: '162',
      solution: [
        'Losing 10% means keeping 90%, so multiply by 0.9 each minute.',
        'After 1 minute: 200 × 0.9 = 180.',
        'After 2 minutes: 180 × 0.9 = 162.',
        'Trap: subtracting 20 twice gives 160. Exponential decay takes 10% of the NEW amount each time.',
      ],
    },
    watch: '15% growth means multiply by 1.15. 10% decay means multiply by 0.90, not 0.10.',
  },
  {
    num: 16, domain: 'M', title: 'Absolute Value', top20: false,
    idea: '|x| is the distance from 0, so it is never negative. |something| = 5 means the something is 5 or −5. Solve both.',
    example: {
      problem: 'Solve |x − 3| = 5.',
      steps: ['Case 1: x − 3 = 5, so x = 8.', 'Case 2: x − 3 = −5, so x = −2.', 'Answer: x = 8 or x = −2.'],
    },
    q1: {
      stem: 'Which values of x satisfy |2x + 1| = 9 ?',
      choices: ['4 only', '4 and −5', '−4 and 5', '4 and 5'], answer: 'B',
      solution: [
        'Case 1: 2x + 1 = 9. Subtract 1: 2x = 8, so x = 4.',
        'Case 2: 2x + 1 = −9. Subtract 1: 2x = −10, so x = −5.',
        'Both work: x = 4 and x = −5.',
      ],
    },
    q2: {
      stem: 'What is the sum of the solutions to |x − 6| = 2 ?',
      answer: '12',
      solution: [
        'Case 1: x − 6 = 2, so x = 8.',
        'Case 2: x − 6 = −2, so x = 4.',
        'Sum: 8 + 4 = 12.',
      ],
    },
    watch: 'If an absolute value equals a negative number, like |x| = −3, there is no solution.',
  },
  {
    num: 17, domain: 'M', title: 'Bonus: Sum and Product of Roots', top20: false, bonus: true,
    formulas: ['For ax^{2} + bx + c = 0:   sum of solutions = −b/a     product of solutions = c/a'],
    idea: 'A shortcut for strong students: you can find the sum or product of the solutions without solving.',
    example: {
      problem: 'For x^{2} − 7x + 10 = 0, find the sum and product of the solutions.',
      steps: ['a = 1, b = −7, c = 10.', 'Sum = −(−7)/1 = 7. Product = 10/1 = 10.', 'Check: the solutions are 2 and 5. 2 + 5 = 7 and 2 × 5 = 10.'],
    },
    q1: {
      stem: 'What is the sum of the solutions to 2x^{2} − 12x + 5 = 0 ?',
      choices: ['−6', '5/2', '6', '12'], answer: 'C',
      solution: [
        'a = 2, b = −12, c = 5.',
        'Sum = −b/a = −(−12)/2 = 12/2 = 6.',
      ],
    },
    q2: {
      stem: 'What is the product of the solutions to 3x^{2} + 4x − 18 = 0 ?',
      answer: '−6',
      solution: [
        'a = 3, b = 4, c = −18.',
        'Product = c/a = −18/3 = −6.',
      ],
    },
    watch: 'The sum has a minus sign in it (−b/a). The product does not (c/a).',
  },

  // ======================== PROBLEM-SOLVING & DATA ========================
  {
    num: 18, domain: 'D', title: 'Ratios and Proportions', top20: true,
    formulas: ['a/b = c/d  →  ad = bc'],
    idea: 'Set up two equal fractions with the same units in the same places, then cross-multiply.',
    example: {
      problem: 'A recipe uses 3 cups of pasta for 4 people. How much pasta for 10 people?',
      steps: ['cups/people: 3/4 = x/10.', 'Cross-multiply: 4x = 30.', 'x = 7.5 cups.'],
    },
    q1: {
      stem: 'In a Minecraft world, the ratio of creepers to zombies is 2 : 5. If there are 35 zombies, how many creepers are there?',
      choices: ['7', '14', '28', '87.5'], answer: 'B',
      solution: [
        'creepers/zombies: 2/5 = c/35.',
        'Cross-multiply: 5c = 70.',
        'c = 14 creepers.',
        'Another way: 35 is 5 × 7, so multiply the 2 by 7 as well: 14.',
      ],
    },
    q2: {
      stem: 'The ratio of boys to girls in a math club is 3 : 4. There are 28 students in the club. How many girls are in the club?',
      answer: '16',
      solution: [
        'The ratio 3 : 4 has 3 + 4 = 7 "parts" in total.',
        '28 students / 7 parts = 4 students per part.',
        'Girls are 4 parts: 4 × 4 = 16 girls (and 3 × 4 = 12 boys).',
        'Check: 12 + 16 = 28, and 12 : 16 simplifies to 3 : 4.',
      ],
    },
    watch: 'A ratio like 3 : 4 is part to part. The whole is 3 + 4 = 7 parts.',
  },
  {
    num: 19, domain: 'D', title: 'Rates and Unit Conversion', top20: true,
    formulas: ['distance = rate × time     Multiply by fractions like (60 minutes / 1 hour) so units cancel'],
    idea: 'Write units on every number. Multiply by conversion fractions so the units you do not want cancel out.',
    example: {
      problem: 'A Roblox player earns 12 Robux every 5 minutes. How many Robux in 2 hours?',
      steps: ['2 hours × (60 minutes / 1 hour) = 120 minutes.', '120 minutes × (12 Robux / 5 minutes) = 1440/5 = 288 Robux.'],
    },
    q1: {
      stem: 'Sam\'s family drives 150 miles to a gaming tournament in 2.5 hours. At the same rate, how many miles would they drive in 4 hours?',
      choices: ['200', '240', '375', '600'], answer: 'B',
      solution: [
        'Rate = distance / time = 150 / 2.5 = 60 miles per hour.',
        'Distance in 4 hours = 60 × 4 = 240 miles.',
      ],
    },
    q2: {
      stem: 'A rectangular soccer field is 20 yards long and 10 yards wide. There are 3 feet in 1 yard. What is the area of the field in square feet?',
      answer: '1800',
      solution: [
        'Convert each side to feet first: 20 yards = 60 feet, 10 yards = 30 feet.',
        'Area = 60 × 30 = 1,800 square feet.',
        'Trap: 20 × 10 = 200 square yards, then × 3 = 600. Wrong! One square yard is 3 × 3 = 9 square feet, so 200 × 9 = 1,800.',
      ],
    },
    watch: 'Square units use the conversion factor squared: 1 square yard = 9 square feet.',
  },
  {
    num: 20, domain: 'D', title: 'Percents and Percent Change', top20: true,
    formulas: ['part = (percent/100) × whole     percent change = (new − old) / old × 100%'],
    idea: 'Percent means "out of 100." For a change, always divide by the ORIGINAL amount.',
    example: {
      problem: 'A $60 Fortnite hoodie is 25% off. What is the sale price?',
      steps: ['25% off means you pay 75%.', '60 × 0.75 = $45.'],
    },
    q1: {
      stem: 'A kids\' song on YouTube had 40,000 views on Monday and 50,000 views on Tuesday. By what percent did the views increase?',
      choices: ['10%', '20%', '25%', '125%'], answer: 'C',
      solution: [
        'Change = new − old = 50,000 − 40,000 = 10,000.',
        'Divide by the OLD amount: 10,000 / 40,000 = 0.25.',
        '0.25 = 25% increase.',
        'Trap: dividing by the new amount gives 20% (choice B).',
      ],
    },
    q2: {
      stem: '30% of what number is 42?',
      choices: ['12.6', '72', '126', '140'], answer: 'D',
      solution: [
        'Translate: 0.30 × n = 42.',
        'Divide by 0.30: n = 42 / 0.3 = 140.',
        'Check: 30% of 140 = 0.3 × 140 = 42. It works.',
        'Trap: 30% of 42 is 12.6 (choice A). That answers a different question.',
      ],
    },
    watch: 'Percent change divides by the ORIGINAL amount, not the new one.',
  },
  {
    num: 21, domain: 'D', title: 'Working Together (Combined Rates)', top20: false,
    formulas: ['rate = 1 / time     combined rate = rate_{1} + rate_{2}     time together = 1 / combined rate'],
    idea: 'If a job takes T hours, the rate is 1/T of the job per hour. Add the rates, then flip to get the time.',
    example: {
      problem: 'Leo paints a wall in 4 hours. Nia paints the same wall in 6 hours. How long together?',
      steps: ['Rates: 1/4 and 1/6 of a wall per hour.', 'Add: 1/4 + 1/6 = 3/12 + 2/12 = 5/12 of a wall per hour.', 'Flip: 12/5 = 2.4 hours (2 hours 24 minutes).'],
    },
    q1: {
      stem: 'Maya can build a Minecraft castle in 3 hours. Theo can build the same castle in 6 hours. How many hours will it take them working together?',
      choices: ['2', '4.5', '9', '1.5'], answer: 'A',
      solution: [
        'Rates: Maya 1/3 castle per hour, Theo 1/6 castle per hour.',
        'Add: 1/3 + 1/6 = 2/6 + 1/6 = 3/6 = 1/2 castle per hour.',
        'Flip: 1 ÷ (1/2) = 2 hours.',
        'Sense check: faster than Maya alone (3 hours). Good.',
      ],
    },
    q2: {
      stem: 'Hose A fills a dunk tank in 10 minutes. Hose A and Hose B together fill it in 6 minutes. How many minutes would Hose B take alone?',
      answer: '15',
      solution: [
        'Together rate = 1/6 tank per minute. Hose A rate = 1/10.',
        'Hose B rate = 1/6 − 1/10 = 5/30 − 3/30 = 2/30 = 1/15 tank per minute.',
        'Flip: Hose B alone takes 15 minutes.',
      ],
    },
    watch: 'Never add the times. Together is always FASTER than the fastest worker alone.',
  },
  {
    num: 22, domain: 'D', title: 'Mean, Median, Mode, Range', top20: true,
    formulas: ['mean = sum / count     sum = mean × count     range = max − min'],
    idea: 'The median is the middle number after you put them in order. An outlier (a very big or very small value) pulls the mean a lot but barely moves the median.',
    example: {
      problem: 'Hot dogs eaten at a picnic: 2, 3, 3, 5, 12. Find the mean, median, mode, and range.',
      steps: ['Mean = (2 + 3 + 3 + 5 + 12) / 5 = 25/5 = 5.', 'Median = middle value = 3. Mode = most common = 3. Range = 12 − 2 = 10.', 'The 12 is an outlier: it pulls the mean (5) above the median (3).'],
    },
    q1: {
      stem: 'Jonah got 4, 6, 7, and 3 eliminations in four Fortnite matches. How many eliminations does he need in the fifth match to have a mean of 6?',
      choices: ['6', '8', '10', '12'], answer: 'C',
      solution: [
        'Use sum = mean × count. For a mean of 6 over 5 matches: 6 × 5 = 30 total.',
        'He has 4 + 6 + 7 + 3 = 20 so far.',
        'He needs 30 − 20 = 10 in the fifth match.',
      ],
    },
    q2: {
      stem: 'A data set is 10, 12, 13, 15, 60. If the value 60 is removed, which statement is true?',
      choices: ['Both the mean and the median decrease, and the mean decreases more.', 'The mean decreases and the median stays the same.', 'Both decrease, and the median decreases more.', 'Neither the mean nor the median changes.'],
      answer: 'A',
      solution: [
        'Before: mean = 110/5 = 22. Median = 13 (the middle of five values).',
        'After removing 60: mean = 50/4 = 12.5. Median = (12 + 13)/2 = 12.5 (four values, so average the two middle ones).',
        'Mean drops by 9.5. Median drops by 0.5.',
        'Both decrease, and the mean decreases much more.',
      ],
    },
    watch: 'Put the numbers in order before finding the median. With an even count, average the two middle numbers.',
  },
  {
    num: 23, domain: 'D', title: 'Probability and Two-Way Tables', top20: true,
    formulas: ['P = favorable / total     P(not A) = 1 − P(A)'],
    idea: 'Probability = what you want / everything possible. The word "given" (or "if a ___ is chosen") shrinks the total to just that row or column.',
    table: {
      caption: 'Favorite food of math club students',
      header: ['', 'Pizza', 'Pasta', 'Hot dogs', 'Total'],
      rows: [['7th grade', '12', '5', '3', '20'], ['8th grade', '10', '8', '2', '20'], ['Total', '22', '13', '5', '40']],
    },
    example: {
      problem: 'Using the table: (a) What is the probability a random student prefers pasta? (b) If an 8th grader is chosen, what is the probability they prefer pasta?',
      steps: ['(a) Pasta total / all students = 13/40.', '(b) Only 8th graders count now: 8/20 = 2/5.'],
    },
    q1: {
      stem: 'Using the table, if a 7th grader is chosen at random, what is the probability that the student prefers pizza?',
      choices: ['12/40', '12/22', '3/5', '22/40'], answer: 'C',
      solution: [
        '"If a 7th grader is chosen" means only the 7th grade row counts. Total = 20.',
        '7th graders who prefer pizza: 12.',
        'P = 12/20 = 3/5.',
      ],
    },
    q2: {
      stem: 'Using the table, if a student who prefers pizza is chosen at random, what is the probability that the student is in 8th grade?',
      answer: '5/11',
      solution: [
        'Now only the Pizza column counts. Total = 22.',
        '8th graders in the Pizza column: 10.',
        'P = 10/22 = 5/11 (about 0.4545). On the test you can type 5/11.',
      ],
    },
    watch: 'Find the "given" group first. That group is your denominator (bottom number).',
  },
  {
    num: 24, domain: 'D', title: 'Scatterplots and Lines of Best Fit', top20: false,
    idea: 'A line of best fit PREDICTS. Its slope is the predicted change in y for each 1-unit increase in x. Actual minus predicted tells you how far off the prediction was.',
    example: {
      problem: 'A line of best fit for practice hours (x) and Fortnite wins (y) is y = 0.8x + 2. Predict wins for 10 hours of practice.',
      steps: ['y = 0.8(10) + 2 = 8 + 2 = 10 wins (predicted).', 'Slope 0.8: each extra hour of practice goes with about 0.8 more wins.'],
    },
    q1: {
      stem: 'A line of best fit for a pasta restaurant is y = 4.5x + 20, where x is the number of hours the restaurant is open in a day and y is the number of pasta dishes sold. What does 4.5 mean?',
      choices: ['The restaurant sells 4.5 dishes in total.', 'For each additional hour open, the model predicts 4.5 more dishes sold.', 'The restaurant is open 4.5 hours a day.', 'The model predicts 20 dishes are sold each hour.'],
      answer: 'B',
      solution: [
        'The slope (the number multiplied by x) is the predicted change in y for each 1 more x.',
        'x is hours open and y is dishes, so 4.5 is the predicted number of extra dishes for each extra hour open.',
      ],
    },
    q2: {
      stem: 'Using the model in Question 1, on a day the restaurant was open 8 hours it actually sold 62 dishes. How many more dishes did it sell than the model predicted?',
      answer: '6',
      solution: [
        'Predicted: y = 4.5(8) + 20 = 36 + 20 = 56 dishes.',
        'Actual minus predicted: 62 − 56 = 6 more dishes.',
      ],
    },
    watch: 'The model predicts; it does not guarantee. Real data points sit above and below the line.',
  },

  // ======================== GEOMETRY ========================
  {
    num: 25, domain: 'G', title: 'Angles', top20: true,
    formulas: ['Straight line 180°   Triangle 180°   Around a point 360°   Polygon (n − 2) × 180°'],
    idea: 'Vertical angles (across from each other where two lines cross) are equal. When parallel lines are crossed by another line, angles in matching positions are equal.',
    example: {
      problem: 'A triangle has angles of 50° and 65°. What is the third angle?',
      steps: ['180 − 50 − 65 = 65°.', 'Two angles are equal, so it is an isosceles triangle.'],
    },
    q1: {
      stem: 'The angles of a triangle measure x°, 2x°, and 3x°. What is the measure of the largest angle?',
      choices: ['30°', '60°', '90°', '120°'], answer: 'C',
      solution: [
        'Angles in a triangle add to 180: x + 2x + 3x = 180.',
        '6x = 180, so x = 30.',
        'The largest angle is 3x = 3(30) = 90°.',
        'Trap: stopping at x = 30 (choice A).',
      ],
    },
    q2: {
      stem: 'A regular hexagon has 6 equal sides and 6 equal angles (like a honeycomb cell). What is the measure, in degrees, of each interior angle?',
      answer: '120',
      solution: [
        'Sum of interior angles = (n − 2) × 180 = (6 − 2) × 180 = 720°.',
        'All 6 angles are equal: 720 / 6 = 120°.',
      ],
    },
    watch: 'Triangles add to 180°, not 360°. Four-sided shapes add to 360°.',
  },
  {
    num: 26, domain: 'G', title: 'Pythagorean Theorem and Special Right Triangles', top20: true,
    formulas: ['a^{2} + b^{2} = c^{2}  (c is the longest side, across from the right angle)', '45-45-90: x, x, x√2     30-60-90: x, x√3, 2x'],
    idea: 'Works only in right triangles. Memorize the common triples (3-4-5, 5-12-13) to save time.',
    example: {
      problem: 'A right triangle has legs 6 and 8. How long is the hypotenuse?',
      steps: ['6^{2} + 8^{2} = 36 + 64 = 100.', 'c = √100 = 10. (It is a 3-4-5 triangle doubled.)'],
    },
    q1: {
      stem: 'A 13-foot ladder leans against a treehouse. The bottom of the ladder is 5 feet from the tree. How high up the tree does the ladder reach?',
      choices: ['8 feet', '12 feet', '14 feet', '18 feet'], answer: 'B',
      solution: [
        'The ladder is the hypotenuse (c = 13). The ground distance is one leg (a = 5).',
        '5^{2} + b^{2} = 13^{2}, so 25 + b^{2} = 169.',
        'b^{2} = 144, so b = 12 feet.',
        'It is the 5-12-13 triple.',
      ],
    },
    q2: {
      stem: 'A square Minecraft farm has a diagonal that is 10√2 blocks long. What is the area of the farm, in square blocks?',
      answer: '100',
      solution: [
        'A diagonal cuts a square into two 45-45-90 triangles. The sides are x, x, x√2.',
        'The diagonal is x√2 = 10√2, so each side x = 10.',
        'Area = 10 × 10 = 100 square blocks.',
      ],
    },
    watch: 'c is always the longest side. To find a leg, SUBTRACT: b^{2} = c^{2} − a^{2}.',
  },
  {
    num: 27, domain: 'G', title: 'Area and Perimeter (including Circles)', top20: true,
    formulas: ['Rectangle A = lw   Triangle A = ½bh   Circle A = πr^{2}, C = 2πr'],
    idea: 'Area covers the inside (square units). Perimeter goes around the outside. For circles, the radius is half the diameter.',
    example: {
      problem: 'A 14-inch pizza (14 inches across) has what area?',
      steps: ['Diameter 14, so radius r = 7.', 'A = π(7)^{2} = 49π, about 153.9 square inches.'],
    },
    q1: {
      stem: 'Which has more pizza: one 16-inch pizza, or two 10-inch pizzas? How much more area does the bigger amount have?',
      choices: ['6π', '14π', '36π', '64π'], answer: 'B',
      solution: [
        '16-inch pizza: radius 8, area = π(8)^{2} = 64π.',
        'Two 10-inch pizzas: radius 5 each, area = 2 × π(5)^{2} = 2 × 25π = 50π.',
        'One 16-inch pizza has more: 64π − 50π = 14π more square inches.',
        'Surprise: one big pizza beats two medium ones.',
      ],
    },
    q2: {
      stem: 'Mia uses 30 blocks of fence to make a rectangular pen in Minecraft. The length is twice the width. What is the area of the pen, in square blocks?',
      answer: '50',
      solution: [
        'Let the width be w. The length is 2w.',
        'Perimeter: 2(length + width) = 30, so 2(2w + w) = 30.',
        '6w = 30, so w = 5 and length = 10.',
        'Area = 10 × 5 = 50 square blocks.',
      ],
    },
    watch: 'Pizza sizes are diameters. Cut in half to get the radius before using πr^{2}.',
  },
  {
    num: 28, domain: 'G', title: 'Volume', top20: false,
    formulas: ['Box V = lwh   Cylinder V = πr^{2}h   (these are on the test reference sheet)'],
    idea: 'Volume is how many unit cubes fit inside. In Minecraft terms: how many blocks fill the space.',
    example: {
      problem: 'A Minecraft room is 5 blocks long, 4 blocks wide, and 3 blocks tall. How many blocks fill it?',
      steps: ['V = 5 × 4 × 3 = 60 blocks.'],
    },
    q1: {
      stem: 'A cylinder-shaped pasta pot has a radius of 5 inches and a height of 8 inches. What is its volume?',
      choices: ['40π', '80π', '200π', '400π'], answer: 'C',
      solution: [
        'V = πr^{2}h = π(5)^{2}(8).',
        '5^{2} = 25, and 25 × 8 = 200.',
        'V = 200π cubic inches.',
      ],
    },
    q2: {
      stem: 'A cube has a volume of 64 cubic units. If every edge is doubled, what is the new volume?',
      choices: ['128', '256', '512', '1,024'], answer: 'C',
      solution: [
        'Find the edge: what number cubed is 64? 4 × 4 × 4 = 64, so the edge is 4.',
        'Double it: new edge = 8.',
        'New volume = 8 × 8 × 8 = 512.',
        'Shortcut: doubling every edge multiplies volume by 2 × 2 × 2 = 8. 64 × 8 = 512.',
      ],
    },
    watch: 'Doubling the edges does not double the volume. It multiplies it by 8.',
  },
  {
    num: 29, domain: 'G', title: 'Similar Triangles', top20: false,
    formulas: ['Lengths × k     Areas × k^{2}'],
    idea: 'Similar shapes have the same angles but different sizes. Matching sides are in the same ratio (the scale factor k).',
    example: {
      problem: 'A 6-foot person casts a 4-foot shadow. At the same time, a tree casts a 20-foot shadow. How tall is the tree?',
      steps: ['height/shadow: 6/4 = h/20.', 'Cross-multiply: 4h = 120.', 'h = 30 feet.'],
    },
    q1: {
      stem: 'Triangle ABC is similar to triangle DEF, with A matching D, B matching E, and C matching F. AB = 4, DE = 10, and BC = 6. What is EF?',
      choices: ['8', '12', '15', '24'], answer: 'C',
      solution: [
        'Scale factor from ABC to DEF: DE / AB = 10/4 = 2.5.',
        'BC matches EF, so EF = 6 × 2.5 = 15.',
      ],
    },
    q2: {
      stem: 'Two triangles are similar. Every side of the big one is 3 times as long as the matching side of the small one. The small triangle has an area of 5. What is the area of the big triangle?',
      answer: '45',
      solution: [
        'Lengths are multiplied by k = 3.',
        'Areas are multiplied by k^{2} = 3^{2} = 9.',
        'Big area = 5 × 9 = 45.',
        'Trap: 5 × 3 = 15. Area grows by the square of the scale factor.',
      ],
    },
    watch: 'Match sides in the same order as the letters (AB with DE, BC with EF).',
  },
  {
    num: 30, domain: 'G', title: 'Coordinate Geometry: Distance, Midpoint, Circles', top20: false,
    formulas: ['Midpoint ((x_{1} + x_{2})/2, (y_{1} + y_{2})/2)     Circle (x − h)^{2} + (y − k)^{2} = r^{2}'],
    idea: 'Midpoint = average the x\'s, average the y\'s. Distance is just the Pythagorean theorem on a grid.',
    example: {
      problem: 'Find the midpoint and the distance between (1, 2) and (4, 6).',
      steps: ['Midpoint: ((1 + 4)/2, (2 + 6)/2) = (2.5, 4).', 'Distance: across 3, up 4, so √(3^{2} + 4^{2}) = √25 = 5.'],
    },
    q1: {
      stem: 'A circle has the equation (x − 3)^{2} + (y + 2)^{2} = 49. What are its center and radius?',
      choices: ['center (3, −2), radius 7', 'center (−3, 2), radius 7', 'center (3, −2), radius 49', 'center (−3, 2), radius 49'], answer: 'A',
      solution: [
        'Match to (x − h)^{2} + (y − k)^{2} = r^{2}.',
        'x − 3 means h = 3. y + 2 is y − (−2), so k = −2. Center (3, −2).',
        'r^{2} = 49, so r = 7.',
      ],
    },
    q2: {
      stem: 'M(5, 1) is the midpoint of segment AB. Point A is (2, −3). What is the sum of the coordinates of point B?',
      answer: '13',
      solution: [
        'Let B = (x, y). Midpoint x: (2 + x)/2 = 5, so 2 + x = 10 and x = 8.',
        'Midpoint y: (−3 + y)/2 = 1, so −3 + y = 2 and y = 5.',
        'B = (8, 5). Sum: 8 + 5 = 13.',
        'Shortcut: A to M goes +3 right and +4 up. Do it again from M: (5 + 3, 1 + 4) = (8, 5).',
      ],
    },
    watch: 'In a circle equation the signs flip for the center, and the number on the right is r^{2}, not r.',
  },
];

// ---------------------------------------------------------------------------
// SPOT THE MISTAKE: one per domain. Shown at the end of each domain's topics.
// "work" lines are the fictional student's (wrong) work. "fix" goes in the key.
// ---------------------------------------------------------------------------
const MISTAKES = {
  A: {
    who: 'Theo',
    task: 'Solve −3x + 4 > 19.',
    work: ['−3x + 4 > 19', '−3x > 15', 'x > −5'],
    fix: [
      'The mistake is in the last line. Theo divided by −3 but did not flip the sign.',
      'Correct: −3x > 15, divide by −3 and flip: x < −5.',
      'Check with x = −6 (which is less than −5): −3(−6) + 4 = 22, and 22 > 19. It works.',
      'Theo\'s answer would allow x = 0: −3(0) + 4 = 4, and 4 is not greater than 19. So his answer is wrong.',
    ],
  },
  M: {
    who: 'Priya',
    task: 'Expand (x + 4)^{2}.',
    work: ['(x + 4)^{2}', '= x^{2} + 4^{2}', '= x^{2} + 16'],
    fix: [
      'The mistake is in the second line. You cannot square each piece separately.',
      'Correct: (x + 4)(x + 4) = x^{2} + 4x + 4x + 16 = x^{2} + 8x + 16.',
      'Check with x = 1: (1 + 4)^{2} = 25. Priya\'s answer gives 1 + 16 = 17. The correct answer gives 1 + 8 + 16 = 25.',
    ],
  },
  D: {
    who: 'Jonah',
    task: 'A $50 Roblox gift card goes up in price by 20%, then the new price goes down by 20%. What is the final price?',
    work: ['Up 20% and down 20% cancel out.', 'Final price = $50'],
    fix: [
      'The 20% decrease is taken from the NEW, bigger price, so the two changes do not cancel.',
      'Up 20%: 50 × 1.20 = $60.',
      'Down 20%: 60 × 0.80 = $48.',
      'Final price: $48, which is 4% less than where it started.',
    ],
  },
  G: {
    who: 'Lena',
    task: 'How many times more pizza is in a 12-inch pizza than in a 6-inch pizza?',
    work: ['12 is twice 6.', 'So the 12-inch pizza has 2 times as much pizza.'],
    fix: [
      'Lena compared the diameters, but pizza is AREA, and area depends on the radius squared.',
      '12-inch: r = 6, A = 36π. 6-inch: r = 3, A = 9π.',
      '36π ÷ 9π = 4. The 12-inch pizza has 4 times as much pizza.',
      'Rule: when lengths double, area multiplies by 2^{2} = 4.',
    ],
  },
};

// ---------------------------------------------------------------------------
// JOKE DECODERS: one per domain. Each "slot" points at a practice question.
// Students write the word that matches their answer. Correct answers spell the
// punchline. The "decoys" are COMMON WRONG ANSWERS, so a mistake produces
// a silly sentence and tells the student to go back and check.
// verify_answers.py checks that each slot's value matches the real answer.
// ---------------------------------------------------------------------------
const JOKES = {
  A: {
    setup: 'Why did the two lines never meet up for pizza?',
    slots: [
      { ref: '1', q: 'q2', value: '10', word: 'THEY' },
      { ref: '3', q: 'q2', value: '1', word: 'HAD' },
      { ref: '4', q: 'q2', value: '4', word: 'THE' },
      { ref: '5', q: 'q2', value: '7', word: 'SAME' },
      { ref: '6', q: 'q2', value: '2', word: 'SLOPE' },
    ],
    decoys: [['3', 'BANANA'], ['−3', 'TACO'], ['5', 'FORGOT'], ['−5', 'NOODLES'], ['8', 'CREEPER'], ['12', 'SOCKS']],
    punchline: 'They had the same slope. (Parallel lines: same slope, never meet, no solution.)',
  },
  M: {
    setup: 'Why did the parabola quit the Fortnite tournament?',
    slots: [
      { ref: '8', q: 'q2', value: '5', word: 'IT' },
      { ref: '10', q: 'q2', value: '46', word: 'ALREADY' },
      { ref: '12', q: 'q2', value: '3', word: 'HIT' },
      { ref: '13', q: 'q2', value: '25', word: 'ITS' },
      { ref: '14', q: 'q2', value: '−13', word: 'PEAK' },
    ],
    decoys: [['50', 'LLAMA'], ['10', 'SPAGHETTI'], ['−1.5', 'DANCED'], ['100', 'BUS'], ['4', 'NAP'], ['−4', 'SNEEZE']],
    punchline: 'It already hit its peak. (A parabola that opens down has a maximum at its vertex.)',
  },
  D: {
    setup: 'Why is the median the calmest kid in math club?',
    slots: [
      { ref: '18', q: 'q2', value: '16', word: 'OUTLIERS' },
      { ref: '19', q: 'q2', value: '1800', word: 'NEVER' },
      { ref: '21', q: 'q2', value: '15', word: 'BOTHER' },
      { ref: '24', q: 'q2', value: '6', word: 'IT' },
    ],
    decoys: [['12', 'PIZZA'], ['600', 'ALWAYS'], ['200', 'TICKLE'], ['4', 'HOT DOGS'], ['56', 'LOVE'], ['62', 'YOUTUBE']],
    punchline: 'Outliers never bother it. (One weird value barely moves the median.)',
  },
  G: {
    setup: 'What do you get when you divide a pizza\'s circumference by its diameter?',
    slots: [
      { ref: '25', q: 'q2', value: '120', word: 'YOU' },
      { ref: '26', q: 'q2', value: '100', word: 'GET' },
      { ref: '27', q: 'q2', value: '50', word: 'PIZZA' },
      { ref: '30', q: 'q2', value: '13', word: 'PI' },
    ],
    decoys: [['720', 'MEATBALL'], ['10', 'SOGGY'], ['15', 'CRUST'], ['30', 'A'], ['3', 'FORK'], ['8', 'PINEAPPLE']],
    punchline: 'You get pizza pi. (Circumference ÷ diameter = π for every circle.)',
  },
};

module.exports = { DOMAINS, FORMULA_SHEET, NOT_ON_TEST, TOPICS, MISTAKES, JOKES };
