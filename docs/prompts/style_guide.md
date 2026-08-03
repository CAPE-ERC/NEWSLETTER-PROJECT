# CAPE Bulletin Style Guide

Derived from the four available editions (March, April, May, July 2026). This
is a starting profile, not a finished one — confirm/extend once the full
2023-to-date archive lands (docs/architecture.md, "Why 2023-to-date data
matters"). Until then, treat `data/raw/bulletins/` as few-shot exemplars
alongside these rules.

## Voice and register

- Formal, analytical, third person. No first-person "we" except in the
  EPU-index attribution line ("Our Nigerian Economic Policy Uncertainty (EPU)
  Index points to...", July 2026).
- Long, clause-dense sentences (typically 30-60 words) that chain cause and
  effect: "X, driven by Y, which in turn Z." Avoid short punchy sentences —
  they don't appear in the source material.
- Consistently hedged, non-committal on direction: "cautiously optimistic,"
  "notwithstanding," "broadly," "albeit," "nonetheless," "however." Almost
  every paragraph pairs a positive development with an immediate caveat.

## Recurring structural moves

- **Global Economic Update** always opens by characterizing the *momentum* of
  the global economy in one sentence ("moderate but increasingly fragile
  momentum," "steady momentum") before breaking into AE / EMDE / China.
- Each regional sub-point follows the same shape: state the trend, name the
  supporting factor, then flag the offsetting risk.
- **Price Update** always reports headline YoY, then MoM, then a food/core
  breakdown in that fixed order, and always closes with a one-sentence
  synthesis ("Overall, ... indicates that...").
- **Fiscal Operations Update** leads with the total FAAC figure and its
  change vs. the prior month, then breaks down Federal/State/LG/13%
  derivation shares, then notes which revenue lines drove the change.
- **Conclusion** always has the same three-paragraph arc: (1) global outlook
  in one paragraph, (2) Nigeria's relative resilience within that global
  context, (3) forward-looking policy priorities. It restates points already
  made in the body rather than introducing new information.
- **Country in Focus** is a single heading line ("Country in Focus –
  [Country]") followed immediately by a figure reference — almost no prose
  in the extracted text itself, most content lives in the referenced chart.

## Figure/number conventions

- Every stat is immediately paired with its prior-period comparator: "X per
  cent in [month], compared with Y per cent in [prior month]."
- Percentage-point changes are called out explicitly for headline inflation
  ("a marginal decline of 0.05 percentage points").
- Currency in Naira uses "₦" with "trillion"/"billion" spelled out, not
  abbreviated (₦153.3 trillion, not ₦153.3tn).
- Figures/tables are referenced inline as "Figure N: <caption>" immediately
  followed by "Source: <org>" on its own line.

## Known inconsistencies to watch for (see data/trend_db/data_notes.md)

- Section content isn't perfectly fixed month to month — March 2026's
  "Fiscal Operations Update" covered public debt (DMO) instead of FAAC.
- "Country in Focus" doesn't strictly rotate — Burkina Faso repeated in
  April and May 2026.
- EPU index numeric values are stated explicitly in some editions and only
  described narratively in others.

## Recurring phrases (verbatim, appear in 2+ editions)

- "cautiously optimistic"
- "broadly resilient" / "underlying resilience"
- "notwithstanding"
- "broad-based" (growth, moderation)
- "data dependent" / "data-dependent approach" (monetary policy)
- "structural rigidities" / "structural bottlenecks"
- "elevated geopolitical uncertainty" / "heightened geopolitical tensions"
