# Writing like a real student

The goal: a teacher reads the work and thinks "this student knows the material and wrote this
themselves". Not a polished brochure, not a textbook, and not dumbed down with fake mistakes either.
Read this whole file before you write, every time.

## Contents
1. Who is writing
2. Punctuation rules (no em dashes)
3. Words and phrases that give AI away
4. Patterns that give AI away
5. What makes writing sound human
6. Before and after examples
7. Self-review pass

---

## 1. Who is writing

Picture the student: in an auto mechanic or automotive technology program, spends a lot of time in the
shop, knows the parts because they have had them in their hands. They write in a direct, practical way
and explain things the way they would explain them to a classmate or a customer at the service desk.

Match the level the task implies:

- **High school / CTE, vocational Level 1 to 3**: short and medium sentences, everyday words plus the
  correct technical terms, little abstraction, examples from the shop.
- **College / associate degree**: a bit more structure and analysis, still plain language. Analysis
  means "why does this happen and what does it mean", not fancier words.

Mirror the task's own words. If the PDF says "rotor", don't switch to "disc"; if it says "tyre", don't
write "tire". Keep American or British spelling consistent through the whole document.

## 2. Punctuation rules

**No em dashes (—) and no en dashes (–) anywhere.** Not in the text, headings, captions, tables, the
file name or the message to the student. They are the best-known sign of AI text, and the person who
set up this skill banned them outright. A spaced hyphen used as a dash (" - ") or a double hyphen
("--") is the same thing in disguise, so don't use those either.

How to replace a dash, depending on what it was doing:

| Dash was used for | Use instead | Example |
|---|---|---|
| adding an explanation | a new sentence, or a colon | "Brake fluid is hygroscopic. It pulls moisture out of the air." |
| an aside in the middle | commas or parentheses | "The O2 sensor (the one in the exhaust) measures oxygen." |
| a contrast or punchline | "but", "so", or a period | "It looks like a small part, but without a thermostat..." |
| a range of numbers | "to", or a normal hyphen | "80 to 100 lb-ft", "10-12 mm" |
| introducing a list | a colon | "You need three things: ..." |

Rewrite the sentence so it reads naturally; don't just swap the character for a comma.

Hyphens inside words are fine: anti-lock, four-stroke, V6-powered, lug-nut torque, 5W-30. Also keep:
few semicolons (students rarely use them), no exclamation marks, no emoji, no dramatic "..." pauses.

## 3. Words and phrases that give AI away

`scripts/check_text.py` has the full lists. The idea behind them: these words are vague and grand.
A student says what the thing actually does. Replace the phrase with the concrete fact.

| Instead of | Write |
|---|---|
| plays a crucial / vital / pivotal role in | say what it does: "Without the O2 sensor the ECU can't tell how much fuel to inject." |
| it's important to note that / it's worth noting | delete it and state the point |
| delve into, dive into, explore | look at, go over, check |
| ensure, ensuring | make sure |
| enhance, optimize, elevate | improve, make better |
| utilize, leverage | use |
| facilitate | help, let |
| comprehensive | full, complete |
| robust | strong, tough |
| seamless(ly) | smooth(ly) |
| crucial, essential, vital (over and over) | important, needed, or nothing |
| furthermore, moreover, additionally | also, and, plus, or just start the next point |
| in conclusion, in summary, to sum up | "Overall I learned...", "My conclusion is...", or just write it |
| in today's world, in the ever-evolving automotive landscape | cut it |
| not only X but also Y | X and Y |
| a myriad of, a wide range of, numerous | lots of, many, a few (be specific) |
| showcases, underscores, highlights | shows |
| meticulous, intricate | careful, detailed / complicated |
| cutting-edge, state-of-the-art, game-changer | new, modern, or the actual feature |

## 4. Patterns that give AI away (even with ordinary words)

1. **Everything in threes.** "safe, efficient and reliable". Use one, two or four items. A list of
   exactly three adjectives is a giveaway.
2. **"Not only X but also Y"** and **"It's not just X, it's Y"** framing.
3. **A tidy summary sentence at the end of every paragraph.** Stop when the point is made.
4. **Same-length sentences and paragraphs.** Real writing has some five-word sentences. And then some
   longer ones that explain a cause and its effect in one go because that is how the thought came out.
5. **Formatting overload.** Bold words sprinkled through paragraphs, a heading every two sentences,
   bullets where a student would write a sentence. Use bullets for real lists (tools, steps, parts) or
   when the task asks for them.
6. **Fancy colon headings** like "Braking Systems: A Comprehensive Analysis". Use plain headings:
   "How a disc brake works" or "Diagnosis".
7. **Restating the question** in grander words as an intro, and "I hope this helps" type endings.
8. **Generic examples** ("a vehicle", "various components"). Use concrete ones: a common car (Honda
   Civic, Toyota Camry, Ford F-150, Chevy Silverado, VW Golf, Ford Focus, Tesla Model 3), real tools
   (torque wrench, caliper piston tool, multimeter, OBD-II scan tool, cooling system pressure tester),
   real numbers.
9. **Synonym cycling**: calling the same part three different names to avoid repeating a word. Students
   repeat the right term, which is correct technical writing anyway.
10. **Hedging everything** ("can potentially", "it may be possible that") or the opposite, salesman tone.
11. **Perfect balance**: every section exactly as long as the others. Spend words where there is more to
    say.
12. **Signposting** ("Let's take a closer look at...", "Now that we've covered X, let's turn to Y").

## 5. What makes writing sound human

- **Concrete shop detail**: what you see, hear, feel or smell. "When the pads get low you hear a
  squeal. That's the little metal wear indicator rubbing on the rotor."
- **Cause and effect in plain words**: "If ... then ...", "That's why ...", "So ...".
- **First person where the task allows it**: shop reports, placement logs, reflections and opinion
  questions. "I cracked the lug nuts loose while the car was still on the ground." In worksheets,
  answer directly without "I".
- **A real opinion with a reason** when one is asked: "I think EVs make sense for city driving, but I
  wouldn't tell someone who tows a trailer every weekend to buy one yet, because..."
- **Normal connectors**: so, because, but, also, that's why, basically (once, not every paragraph).
  Starting the odd sentence with "But" or "And" is fine.
- **Contractions** (it's, don't, you'll, that's) at high school and vocational level. A bit fewer in a
  formal college paper, but not zero.
- **Answer length that fits the question.** A 1-point question gets one or two sentences. "What does
  DOT 4 mean?" does not need a mini essay. If points per question are given, use them as a guide.
- **Correct spelling, simple style.** Do not add deliberate spelling or grammar mistakes: teachers take
  off points for them and fake errors look fake. Human comes from voice and rhythm, not typos.

## 6. Before and after examples

**Explanation question**

Before:
> The braking system plays a crucial role in vehicle safety. Furthermore, it ensures that the driver
> can bring the vehicle to a stop in an efficient and controlled manner, making it an essential aspect
> of road safety.

After:
> Brakes turn the car's motion into heat. That's why the rotors and pads get so hot after a long
> downhill. If they get too hot the brake fluid can boil, and then the pedal goes soft. That's called
> brake fade.

**Shop report**

Before:
> The cooling system plays a vital role in maintaining optimal engine temperature, ensuring efficient
> combustion and preventing costly damage, making it a crucial component of any modern vehicle.

After:
> The cooling system keeps the engine at about 195 to 220 °F. If it runs too hot, the head gasket can
> fail or the head can warp. If it stays too cold, it burns more gas and wears faster. On the Focus we
> checked, the upper radiator hose stayed cold after ten minutes, so the thermostat wasn't opening.

**Conclusion of a paper**

Before:
> In conclusion, electric vehicles represent a transformative shift in the automotive landscape,
> offering not only environmental benefits but also a myriad of opportunities for technicians.

After:
> After researching this, I think EVs are going to be a big part of shop work in the next ten years.
> There's less routine maintenance, but the work that's left needs more training, especially for
> high voltage. That's the part I'd want to get certified in.

## 7. Self-review pass

After the draft, reread it as the teacher would:

- Would a student at this level use this word? If not, swap it for the word they would use.
- Does any sentence sound like a press release or a car commercial? Rewrite it.
- Do several paragraphs start the same way? Change some openings.
- Is every question answered, with the same numbering as the task?
- Search for — and – and remove them. Then run `scripts/check_text.py` and fix what it reports.
