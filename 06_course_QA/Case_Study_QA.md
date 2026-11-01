# ECON3700 Case Study Q&As

**Course:** [ECON3700] Environmental Economics (St Lucia), Semester 2, 2026 (ECON3700_7660_60564)
**Forum:** Discussions > Case Study Q&As
**Created by:** Ammarr Aripin (Instructor)

**Forum rules:**
- Post questions for clarification only.
- Review existing questions and answers before posting. Questions already addressed will not receive a further response.
- Queries posted within 24 hours of the submission deadline will not be answered.

Threads are listed newest first, as on the forum.

---

## 1. Significant digits (Yuxuan Zhou, 29 Oct 2026)

**Q (Yuxuan Zhou):** Should we consider the number of significant digits for different data?

**A (Milly Coburn, 29 Oct 2026):**
- Keep the inputs as provided and retain full precision throughout spreadsheet calculations.
- Do not round intermediate results, as this can affect subsequent calculations.
- For presentation, where no precision is specified, use a sensible and consistent level of rounding appropriate to the units and underlying data.
- Use Excel cell formatting to display rounded values while retaining full precision.

**Follow-up (Milly Coburn, 30 Oct 2026):**
- The value given in the provided Excel is 10.005. It appears as 10.0 in the screenshot only because the student's Excel is displaying one decimal place.

---

## 2. Vegetation management schedule, Scenarios 3 and 4 (Baojia Yuan, 28 Oct 2026)

**Q (Baojia Yuan):**
- For Scenario 4, since vegetation management begins in FY2027, should FY2027 use the first-year rates ($1,500/km grass spraying and $2,705/km woody weed control), with subsequent-year rates from FY2028?
- For Scenario 3, should vegetation management begin in FY2026 or FY2027? Should first-year rates apply in whichever year it begins, followed by subsequent-year rates?

**A (Milly Coburn, 28 Oct 2026):**
- Scenario 4: apply first-year vegetation-management rates in FY2027, then the subsequent-year schedule from FY2028 onwards.
- Scenario 3: vegetation management also begins in FY2027. Apply first-year rates in that year, followed by the subsequent-year rates specified in the inputs.

---

## 3. Remaining spreadsheet clarifications (Joshua Sommer, 26 Oct 2026)

**Q (Joshua Sommer):**
- Doshi et al. (2026, Section 2.4.2) gives Scenario 4 the same vegetation management costs as Scenarios 2 and 3. Should rows 57-58 be included, and should Year 1 rates start in FY2027 alongside other ongoing Scenario 4 costs?
- For Scenario 4 in FY2026, should sediment delivery use the business-as-usual erosion rate (55.3 t/km/yr), given the reduced rate starts in FY2027?
- For Scenario 3, do the reduced erosion rate and 70% maintenance begin in FY2026 (same year as capital works) or FY2027 as with Scenario 4?
- Is the avoided abatement cost in Inputs!C40 the $1,624/t/yr (FY2024 dollars) used in Doshi et al. (2026)?
- For CPI indexation, should we use financial-year average CPI or June-quarter CPI?
- Inputs!C5 (=C4+30) returns FY2056, while the DCF runs FY2026-FY2055. Should this read FY2055?

**A (Milly Coburn, 27 Oct 2026, edited):**
- Several of these points have already been addressed in previous posts.
- Scenario 3 timing: distinguish the initial works year from subsequent operating years. The reduced maintenance and erosion assumptions apply from FY2027.
- Abatement value: we do not confirm whether selected input values are correct on this forum.
- Assessment end year: use the 30 annual periods shown in the DCF, FY2026-FY2055 inclusive.

---

## 4. Scenario 4 vegetation management in DCF rows 57 and 58 (Baojia Yuan, 26 Oct 2026)

**Q (Baojia Yuan):**
- Should Scenario 4 include invasive grader grass spraying and woody weed control costs in DCF rows 57 and 58? The Inputs notes list these for Scenarios 2 and 3 only, while Section 2.4.2 of Doshi et al. (2026) states Scenario 4 has the same vegetation management costs.
- For this simplified model, should they be included or set to zero? If included, from FY2026 or FY2027?

**A (Milly Coburn, 27 Oct 2026):**
- Include vegetation management within Scenario 4's ongoing costs, using the provided grass-spraying and woody-weed-control inputs.
- These begin in FY2027.

---

## 5. Valuation, Scenario 4 maintenance, CEA, CPI (Angelina Abery, 25 Oct 2026)

**Q (Angelina Abery):**
- Environmental valuation: is it appropriate to use the avoided fine-sediment abatement cost of $488/tonne from the Reef Trust evaluation, or is there a specific valuation recommended? If using $488, should it be indexed to FY2026 dollars, and what base year?
- Scenario 4 maintenance: should FY2026 gravel maintenance use the full business-as-usual rates for both road and drain maintenance, with bitumen maintenance and drain cleaning commencing in FY2027?
- CEA: should cost per tonne of sediment avoided divide total discounted financial costs by discounted cumulative tonnes avoided over 30 years, or by annual tonnes avoided relative to BAU? Should sediment reductions also be discounted at 7%?
- CPI indexation (FY2024 to FY2026): use average CPI for each financial year or June-quarter CPI?

**A (Ammarr Aripin, Instructor, 27 Oct 2026):**
- Your question has already been addressed in a previous post by another student. Please refer to the earlier responses.
- We do not provide confirmation of whether an answer is correct or incorrect on this platform.

---

## 6. CEA basis, Scenario 4 timing, resealing (Molly Jones, 25 Oct 2026)

**Q (Molly Jones):**
- For CEA, should cost per tonne of sediment avoided use each scenario's total discounted financial costs, or incremental discounted financial costs relative to BAU?
- For Scenario 4, should drain cleaning and the reduced sediment rate begin in FY2026 or FY2027?
- Does the simplified cost structure exclude a separate resealing payment after 20 years, despite resealing being mentioned in the presentation?

**A (Milly Coburn, 25 Oct 2026):**
- CEA: use each scenario's total discounted financial costs over the assessment period. Measure the sediment avoided relative to BAU.
- Scenario 4 timing: drain cleaning and the reduced sediment rate begin in FY2027, following the initial works year.
- Resealing: yes, the simplified spreadsheet treatment includes initial sealing expenditure and ongoing maintenance, without a separate resealing payment after 20 years.

---

## 7. Scenario 4 first-year gravel maintenance and reseal (Anderson Ma, 24 Oct 2026)

**Q (Anderson Ma):**
- First-year gravel maintenance (FY2026): full Scenario 1 (BAU) maintenance cost, or the reduced 70% rate used in Scenarios 2 and 3?
- Reseal after 20 years: is there a specific reseal cost, or assume it equals the original bitumen dust seal cost ($110,000/km)?

**A (Milly Coburn, 24 Oct 2026):**
- First-year gravel maintenance: refer to the existing gravel-road maintenance inputs. The reduced rates specified for Scenarios 2 and 3 should not automatically be applied to Scenario 4.
- Resealing: follow the simplified cost structure using the initial sealing expenditure and ongoing annual maintenance costs.

---

## 8. Stream crossings and sediment delivered (Anderson Ma, 23 Oct 2026, edited)

**Q (Anderson Ma):**
- Number of stream crossings treated: 1 crossing per km x crest-to-crest crossing length (10.005 km), or x total road length (21.5 km)?
- Should rock mulch, grade control structures and gully rock chutes (AUD$/crossing) be applied to every stream crossing?
- Sediment delivered to the reef: erosion rate (t/km/yr) x total road length (21.5 km), or x crest-to-crest crossing length (10.005 km)?

**A (Milly Coburn, 23 Oct 2026):**
- Stream crossings treated: multiply the assumed crossing density (1 per km) by the total road length (21.5 km).
- Rock mulch, grade control structures and gully rock chutes: apply the provided per-crossing cost to all assumed stream crossings treated in Scenarios 3 and 4.
- Sediment delivered to the reef: multiply the applicable erosion rate by the crest-to-crest crossing length (10.005 km). Use the erosion rate appropriate to the scenario and year being assessed.

---

## 9. Concrete vs rock floodways (Baojia Yuan, 22 Oct 2026)

**Q (Baojia Yuan):**
- In the Week 12 guest lecture, Scenario 3 uses "rock or concrete bed-level floodways at stream crossings". For capital costs, are concrete and rock floodways alternatives, or should both costs apply to all crossings? If alternatives, how should allocation be determined?

**A (Milly Coburn, 23 Oct 2026):**
- In this simplified model, include both concrete and rock floodway costs as capital-cost components in Scenario 3.
- Apply each per-crossing unit cost to the assumed number of stream crossings treated. No allocation between the two treatments is needed.

---

## 10. Scenario 4 timing, rows 53 and 54 (Molly Jones, 22 Oct 2026)

**Q (Molly Jones):**
- Rows 53 and 54 specify gravel maintenance for "FY2025 only", but the model starts in FY2026. Should we (a) leave these rows at zero and include sealing and bitumen maintenance from FY2026, or (b) include gravel maintenance in the first model year, with sealing and bitumen maintenance starting later?
- Should the first year's annual maintenance costs have a discount factor of 1, or be discounted by one year?

**A (Milly Coburn, 22 Oct 2026, edited 24 Oct 2026):**
- Interpret the "FY2025 only" labels in rows 53-54 as referring to the first model year, FY2026.
- Gravel maintenance applies during that initial year, alongside the sealing and road-preparation expenditure. Ongoing bitumen maintenance then begins in FY2027. You do not need to delay sealing expenditure until the second model year.
- The first year's annual maintenance costs should be discounted by one year, with FY2026 treated as year 1 for discounting.

---

## 11. Starting year for the DCF (Jayden Anders, 21 Oct 2026)

**Q (Jayden Anders):**
- With an assessment period of 2026 to 2055, should FY26 be Year 1 for ongoing costs, with a discount factor consistent with Tutorial 4?
- For Scenarios 3 and 4, should initial capital expenditure in FY26 be treated as a Year 0 cost (not discounted)?

**A (Ammarr Aripin, Instructor, 22 Oct 2026):**
- All values should be discounted starting from FY26 as Year 1, using a discount factor consistent with Tutorial 4.
- Calculate discount factors for FY26 to FY55 in Row 14 using: Discount Factor = 1 / (1+dr)^n, where n is the year number (FY26 = Year 1).
- For Scenarios 3 and 4, think this through more carefully, but the key point is that values are still discounted from FY26 onwards.

---

## 12. CPI indexation factor, FY2024 to FY2026 (Jingyu Chen, 21 Oct 2026)

**Q (Jingyu Chen):**
- Should we use CPI values for June 2024 and June 2026, or the average CPI values for FY2024 and FY2026?
- Can we enter the CPI values and calculate the ratio directly in the assignment spreadsheet, or must we download the ABS CPI spreadsheet and reference it?

**A (Ammarr Aripin, Instructor, 22 Oct 2026):**
- Think about how CPI is used to convert values from FY2024 dollars to FY2026 dollars. Identify the appropriate CPI values for the two financial years and calculate the ratio of the later year's CPI to the earlier year's CPI.
- There is no requirement to link to the ABS CPI workbook. Enter the relevant CPI values directly in your spreadsheet and calculate the factor there. Make sure source CPI values are correct and the calculation is clearly shown.

---

*Forum stats at time of capture: 12 posts, 14 replies.*
