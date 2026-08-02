# Data quality notes — March/April/May/July 2026 editions

Found while manually curating `data/processed/*.json` from the four available
bulletins. Worth checking against the source docs (NBS/FAAC releases) once the
fuller archive lands, and worth the editor agent flagging automatically in
future editions (see docs/architecture.md, Phase 4).

1. **FAAC month labeling is inconsistent across editions.** The April 2026
   edition describes ₦1.894tn as disbursed "for February 2026" revenue. The
   May 2026 edition later describes the same figure (₦1.89tn, rounded) as
   "distributed in March 2026." Both can't be right about the underlying
   collection month — kept as each edition originally labeled it rather than
   silently reconciled.

2. **March 2026's "Fiscal Operations Update" section covers public debt
   (DMO data)**, not FAAC allocations like the other three editions. Not an
   error, just a structural inconsistency in what that section covers
   month to month — worth deciding whether the automated pipeline should
   normalize this to always be FAAC, or accept that the section's content
   varies.

3. **"Country in Focus" repeated Burkina Faso** in both the April and May
   2026 editions, breaking the implied monthly-rotation pattern. Possibly a
   copy-paste carryover in the source document rather than an intentional
   editorial choice.

4. **EPU index numeric values are rare.** Of the four editions, only July
   2026 states explicit index values in the text (132 in April, 137 in May).
   March, April, and May all discuss EPU direction/drivers narratively
   without citing a number — meaning most of this trend series can't be
   backfilled from these four documents alone.

5. **GDP growth figures are only restated in the March 2026 edition**
   (Q4 2025 real GDP, annual 2025 growth). April, May, and July all reference
   GDP-related figures/charts without repeating the numbers in prose, so
   those fields are null in the corresponding processed JSON rather than
   guessed from the chart captions.

6. **The May 2026 bulletin's source .docx has a corrupted embedded image**
   (`word/media/image4.png`, bad CRC-32) — present in the original uploaded
   zip, not introduced during scaffolding. `python-docx`'s `Document()` call
   fails outright on this file; text was recovered via a raw-XML fallback
   that reads `word/document.xml` directly, bypassing the corrupt media part.
