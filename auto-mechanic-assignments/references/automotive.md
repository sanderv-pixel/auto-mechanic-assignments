# Getting the automotive content right

Teachers in a technical program grade the technical content first. A beautifully written answer with a
wrong torque value, a wrong unit or a made-up trouble code loses the points. Use this file for accuracy
rules, formulas, reference values, unit conversions, standard answer structures, safety points and
US/UK terminology.

## Contents
1. Accuracy rules
2. Formulas with worked examples
3. Unit conversions
4. Reference values that come up a lot
5. Answer structures teachers expect
6. Safety and environment points
7. Sources
8. US / UK terminology

---

## 1. Accuracy rules

- **The task's own material wins.** If the user also gave you class notes, a textbook page, a service
  manual extract or a data sheet, take facts, values and terms from there first. If the task refers to
  material you don't have ("see chapter 12"), use standard textbook knowledge in the same terms the
  task uses.
- **Never invent manufacturer-specific data** (exact torque for one model, what a P1xxx code means for
  one brand, service intervals, part numbers, TSBs). If web search is available, look it up and use real
  data. If not, give the typical range and say the exact value comes from the service information
  ("per the manufacturer's specs", "check the service manual").
- **Every number gets a unit**, and the unit system matches the task. American tasks usually use lb-ft,
  psi, °F, inches, mph and quarts; British and international tasks use Nm, bar or kPa, °C, mm, km/h and
  litres. Electrical units are the same everywhere (V, A, Ω, W). Use the symbols directly: π × ² ³ √ Ω
  ° ≈ ≤ ≥.
- **Calculations: formula, substitution, result.** Show it the way the student would on paper: write
  the formula, fill in the numbers with units, give the result with a unit and sensible rounding, then
  one sentence saying what it means if the question asks for interpretation.
- **Measured values in a shop report** must be plausible and consistent with each other (a battery at
  12.4 V is not "fully charged"; a 50/50 coolant mix does not protect to −80 °F). If you had to make up
  values because the student did the practical and you don't have their numbers, tell them in the
  handover message which values to replace with their own measurements.

## 2. Formulas with worked examples

**Engine displacement**
Single cylinder: V = (π / 4) × bore² × stroke. Total = V × number of cylinders.
Bore 81 mm, stroke 95.5 mm: V = 0.7854 × 8.1² × 9.55 cm ≈ 492 cm³; four cylinders ≈ 1968 cm³ (2.0 L).
In inches: bore 4.00 in, stroke 3.48 in, V8: 0.7854 × 16 × 3.48 × 8 ≈ 350 cubic inches (a 350 small block).

**Compression ratio**
CR = (swept volume + clearance volume) / clearance volume.
Swept 492 cm³, clearance 29 cm³: CR = 521 / 29 ≈ 18 : 1. Gas (petrol) engines are usually about 9 to
13 : 1, diesels about 15 to 20 : 1.

**Mean piston speed**
vm = 2 × stroke × rpm / 60 (stroke in m gives m/s). Stroke 0.0955 m at 4000 rpm: vm ≈ 12.7 m/s.

**Power and torque**
US: hp = torque (lb-ft) × rpm / 5252. 300 lb-ft at 4000 rpm: 300 × 4000 / 5252 ≈ 228 hp.
Metric: P (kW) = torque (Nm) × rpm / 9549. 320 Nm at 1750 rpm: ≈ 58.6 kW.
Torque and horsepower curves always cross at 5252 rpm when plotted in lb-ft and hp.

**Electrical**
Ohm's law: V = I × R. Power: P = V × I (= I² × R = V² / R).
Series: Rtotal = R1 + R2 + ...; current is the same everywhere, voltage divides.
Parallel: 1/Rtotal = 1/R1 + 1/R2 + ... (two resistors: R1 × R2 / (R1 + R2)); voltage is the same
across each branch, current divides.
A 55 W headlight bulb at 12 V: I = 55 / 12 ≈ 4.6 A; R = 12 / 4.6 ≈ 2.6 Ω (hot filament).
Voltage drop test: more than about 0.2 V across a connection or about 0.5 V across a whole cable run
under load points to resistance there.
Battery energy: Wh = V × Ah. 12 V, 70 Ah: 840 Wh.

**Hydraulic brakes**
Pressure = force / area, area = (π / 4) × d². Pressure is the same everywhere in the system, so
output force / input force = output piston area / input piston area.
US example: 200 lb on a 1.0 in master cylinder piston: area = 0.785 in², pressure = 200 / 0.785 ≈ 255
psi. Caliper piston 2.0 in: area = 3.14 in², force = 255 × 3.14 ≈ 800 lb (four times the input,
because the area is four times bigger).
Metric example: 1000 N on a 22 mm piston: area ≈ 380 mm², pressure ≈ 2.63 N/mm² = 26.3 bar; caliper
piston 54 mm: area ≈ 2290 mm², force ≈ 6025 N. (1 N/mm² = 1 MPa = 10 bar.)
Pedal ratio: output force = pedal force × (pedal length / pushrod distance), for example 4 : 1 or 6 : 1.

**Gear ratios and speed**
Gear ratio = driven teeth / driving teeth. Output rpm = input rpm / ratio. Output torque ≈ input
torque × ratio (minus losses). Overall ratio = transmission ratio × final drive (axle) ratio.
Vehicle speed (mph) = rpm × tire diameter (in) / (overall ratio × 336).
Tire diameter from size 225/60R16: sidewall = 225 × 0.60 = 135 mm; diameter = 16 × 25.4 + 2 × 135 =
676 mm ≈ 26.6 in.

**Fuel economy**: mpg = miles driven / gallons used; L/100 km = litres used / km driven × 100.
**Coolant mix (drain and fill)**: amount to drain = capacity × (target % − current %) / (100 % −
current %). 7.0 L at 30 %, target 50 %: 7.0 × 20 / 70 = 2.0 L drained and replaced with pure antifreeze.

## 3. Unit conversions

| From | To | Multiply by |
|---|---|---|
| Nm | lb-ft | 0.7376 (1 lb-ft = 1.356 Nm) |
| bar | psi | 14.5 (1 psi = 0.0689 bar = 6.895 kPa) |
| kW | hp | 1.341 (1 hp = 0.7457 kW; 1 PS = 0.7355 kW) |
| km | miles | 0.6214 |
| litres | US gallons | 0.2642 (1 US gal = 3.785 L; 1 UK gal = 4.546 L) |
| litres | US quarts | 1.057 |
| mm | inches | divide by 25.4 |
| °C | °F | × 9/5 + 32 |
| cm³ | cubic inches | divide by 16.387 |

## 4. Reference values that come up a lot

Typical values. When a task gives its own values, use those.

- **Tire size** 225/60R16 98H: width 225 mm, sidewall 60 % of width, R = radial, 16 inch rim, load
  index 98 = 750 kg (1653 lb) per tire, speed rating H = 130 mph / 210 km/h (S 112 mph, T 118, H 130,
  V 149, W 168, Y 186). The last four digits of the DOT code give week and year (2324 = week 23 of
  2024). Minimum tread depth: 2/32 inch in most US states (4/32 inch is the usual advice to replace),
  1.6 mm in the UK and EU. Use the pressure on the door jamb placard, not the number on the sidewall
  (that is the maximum).
- **Brake fluid (FMVSS 116)**: DOT 3, DOT 4 and DOT 5.1 are glycol based and can be mixed; DOT 5 is
  silicone based and must not be mixed with them. Minimum dry / wet boiling points: DOT 3 401 / 284 °F
  (205 / 140 °C), DOT 4 446 / 311 °F (230 / 155 °C), DOT 5.1 500 / 356 °F (260 / 180 °C). It absorbs
  water (hygroscopic), which lowers the boiling point, so it is replaced every 2 to 3 years or per the
  manufacturer.
- **12 V battery, resting voltage**: about 12.6 V full, about 12.4 V ≈ 75 %, about 12.2 V ≈ 50 %,
  below 12.0 V deeply discharged. Charging voltage with the engine running about 13.5 to 14.7 V. During
  cranking it should not drop below about 9.6 V. Disconnect the negative cable first, reconnect it last.
  CCA (cold cranking amps) is the current the battery can deliver at 0 °F for 30 s while staying above
  7.2 V.
- **Cooling**: normal operating temperature about 195 to 220 °F (90 to 104 °C). The thermostat starts
  opening around 180 to 195 °F (82 to 90 °C) depending on the engine. A 15 psi (about 1 bar) pressure
  cap raises the boiling point of 50/50 coolant to about 265 °F (130 °C). A 50/50 mix protects to about
  −34 °F (−37 °C).
- **Engine oil**: 5W-30 = cold viscosity grade 5W (W = winter), hot viscosity grade 30. API service
  categories (SP, SN, CK-4) and OEM approvals (GM dexos, VW 504 00/507 00) matter as much as the
  viscosity.
- **Air/fuel ratio**: stoichiometric is about 14.7 : 1 for gasoline (λ = 1). Narrowband O2 sensor
  switches between about 0.1 V (lean) and 0.9 V (rich) around 0.45 V. Fuel trims beyond about ±10 %
  point to a mixture problem the ECU is compensating for.
- **OBD-II**: required on US cars from 1996 (EOBD in Europe from 2001 for gas, 2004 for diesel).
  First letter P (powertrain), B (body), C (chassis), U (network). P0xxx generic, P1xxx manufacturer
  specific. Examples: P0300 random misfire, P0171 system too lean bank 1, P0420 catalyst efficiency
  below threshold bank 1, P0128 coolant temperature below thermostat regulating temperature. A code
  points to a circuit or a symptom, not straight to a broken part; you still have to test.
- **A/C refrigerant**: R-134a in older cars, R-1234yf in most new cars. In the US, handling refrigerant
  requires EPA Section 609 certification; in the UK, F-gas certification.
- **High voltage (hybrid / EV)**: anything above 60 V DC is high voltage in a vehicle; traction systems
  run at roughly 200 to 800 V. Orange cables are HV. Only trained, qualified people work on it (US: OEM
  training and ASE L3 hybrid/EV specialist; UK: IMI EV levels).
- **Lug nuts / wheel bolts**: passenger cars usually around 80 to 100 lb-ft (110 to 140 Nm), trucks
  more; exact value from the manufacturer. Tighten in a star pattern with a torque wrench, never with
  an impact gun alone.

## 5. Answer structures teachers expect

**"Explain how X works"**
What it is for (why the vehicle needs it), the main parts, how it works step by step, what goes wrong
and how you notice it (symptoms), how you inspect or maintain it.

**Diagnosis or case study** (the "3 Cs": concern, cause, correction)
Customer concern and symptoms, verify the concern, possible causes with the most likely first, which
tests or measurements you do and what values you expect, what you found, the repair, and verification
afterwards (road test, clear codes, retest).

**Work procedure / shop report**
Repair order and preparation (fender covers, seat cover, floor mat), PPE, lifting the vehicle at the
correct lift points, tools and supplies, the steps in the right order with torque specs, the final
check and road test, cleanup and waste disposal, what you tell the customer.

**Comparison**
A table often works best, followed by a short paragraph with your own conclusion.

**Customer explanation**
No jargon, explain the risk if they don't repair it, give a rough cost or time if the task asks for it.

**Reflection**
What went well, what was hard, what you'd do differently, what you learned. First person, honest, short.

## 6. Safety and environment points

Mention these when they fit the task; teachers often give points for them.

- PPE: safety glasses, nitrile gloves, steel-toe boots, hearing protection when needed. No loose
  clothing or jewelry near moving parts.
- Lift the vehicle at the correct lift points; use jack stands, never work under a car held up by a
  jack alone. Chock the wheels.
- Battery: disconnect negative first. No sparks near a charging battery (hydrogen gas).
- Brake dust: don't blow it off with compressed air, use brake cleaner or a wet wash.
- Brake fluid damages paint. Coolant is toxic and tastes sweet, so clean up spills (pets, kids).
- Never open a hot radiator cap or expansion tank: the system is under pressure.
- Used oil, filters, coolant, brake fluid, refrigerant and batteries go to proper recycling (EPA rules
  in the US, hazardous waste rules in the UK).
- Hybrid / EV: power down, remove the service disconnect per the procedure, lock it out so no one can
  switch it back on, verify zero voltage, wait for the capacitors to discharge, use insulated tools and
  rated insulating gloves.
- Running engines indoors: exhaust extraction (carbon monoxide).

## 7. Sources

Only cite sources that exist. Never invent a URL, article, author or page number.

- Material the student uploaded (class notes, textbook pages): cite it by its title.
- If web search is available: use real pages (manufacturer sites, NHTSA, EPA, ASE, government or
  technical sites) and list them with the URL and the date accessed.
- Textbooks: cite one only if the student or the task names it (for example the class textbook), or
  you could check the page. Common ones are James D. Halderman, *Automotive Technology: Principles,
  Diagnosis, and Service* (Pearson); *Modern Automotive Technology* (Goodheart-Willcox); Hillier's
  *Fundamentals of Motor Vehicle Technology* (UK); *Bosch Automotive Handbook* (Bosch / Wiley).
  Prefer sources you actually opened and read.
- If a source title contains a dash, write it with a colon instead ("EV Safety Basics: Part II").
  The no-dash rule covers the reference list too.
- Use the citation style the task asks for (APA, MLA, Harvard). If none is given, use APA 7.

## 8. US / UK terminology

Use the column that matches the task, and stay consistent.

| US | UK |
|---|---|
| gas, gasoline | petrol |
| hood | bonnet |
| trunk | boot |
| fender | wing |
| windshield | windscreen |
| tire | tyre |
| rotor | brake disc |
| parking brake / emergency brake | handbrake |
| transmission | gearbox |
| stick shift / manual | manual gearbox |
| muffler | silencer |
| lug nuts | wheel nuts / wheel bolts |
| wrench | spanner |
| jack stand | axle stand |
| shop | workshop / garage |
| repair order (RO) | job card |
| antifreeze / coolant | coolant / antifreeze |
| odometer | odometer / mileometer |
| state inspection / emissions test | MOT |
| ASE certification | IMI accreditation |
| horsepower (hp) | brake horsepower (bhp) / kW |
| lb-ft, psi, °F, quarts | Nm, bar, °C, litres |
