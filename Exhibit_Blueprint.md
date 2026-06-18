# Exhibit Blueprint — Apparel & Footwear Whitespace → Zudio Store Potential

**How to use this:** build *wide* first. Every exhibit below has a home tab in `Zudio_Whitespace_Engine.xlsx` — once you paste real data, the `CALC_` and `OUT_` tabs populate and these become chartable. Then you cut: keep only the ~22 that carry the spine (flagged **S#** in the last column). The rest stay in an appendix or get killed.

**The workflow, in four passes:**
1. **Build** — generate all ~50 exhibits from the engine (data → CALC → OUT tabs).
2. **Annotate** — write the one-line takeaway *under* each. If you can't, the exhibit is decoration — cut it.
3. **Distil** — keep the exhibits whose takeaways chain into one argument: *big prize → concentrated demand → thin supply down-tier → that gap is the runway → comps prove the method → here's Zudio's number.*
4. **Deck** — the survivors become the 22 slides.

Legend for "Confidence": **H** = hard/observed · **M** = modelled/assumption · **P** = projected.

---

## ACT 1 — THE PRIZE  (how big, how fast)

| # | Exhibit | Form | Source tab | The insight it must land | Conf | Deck |
|---|---------|------|-----------|--------------------------|------|------|
| 1 | India apparel+footwear wallet — headline TAM | Big-number callout | `10_CALC_Demand` total | "The addressable wallet is ₹X lakh Cr today." | P | **S3** |
| 2 | Apparel vs footwear split | Donut / stacked bar | `4_IN_Demand` | Mix and which grows faster | P | — |
| 3 | Wallet trajectory, FY-by-FY (hist + projected) | Column + CAGR line | `10_CALC_Demand` (proj col) | The growth slope and CAGR | P | **S3** |
| 4 | Wallet growth bridge: population × per-capita × share drift | Waterfall | `10_CALC_Demand` | *What* drives growth — most of it is per-capita, not population | P | **S4** |
| 5 | Per-capita spend by tier | Bar | `10_CALC_Demand` | Lower tiers spend a fraction of T1 → headroom as they converge | P | **S5** |
| 6 | Per-capita spend vs per-capita GDP (states) | Scatter + fit | `4_IN_Demand` | Spend tracks income → projection is grounded, not hopeful | P | — |
| 7 | MPCE growth path by geography | Line | `4_IN_Demand` | The single lever the whole projection rests on | P | — |

## ACT 2 — WHERE THE MONEY IS  (demand geography)

| # | Exhibit | Form | Source tab | The insight it must land | Conf | Deck |
|---|---------|------|-----------|--------------------------|------|------|
| 8 | Wallet by tier (apparel/footwear stacked) | Stacked bar | `10_CALC_Demand` | How demand concentrates by tier | P | **S6** |
| 9 | Wallet by population slab (0–5 Mn bands) | Bar | `2_IN_Pincode` | Demand isn't only in mega-cities | P | — |
| 10 | Demand concentration (Lorenz) — cum. wallet vs cum. pincodes | Curve | `2_IN_Pincode` | "Top X% of pincodes = Y% of wallet" — focuses the rollout | P | **S7** |
| 11 | Top 20 cities by apparel wallet | Ranked bar / table | `3_IN_City` | The anchor markets | P | **S8** |
| 12 | Wallet by state | Choropleth / bar | `4_IN_Demand` | Regional skew | P | — |
| 13 | Tier × zone wallet matrix | Heatmap | `10_CALC_Demand` | Geographic pockets where wallet sits | P | — |
| 14 | Rural vs urban wallet — size & growth | Paired bar | `2_IN_Pincode` | The rural/low-tier growth thesis | P | — |

## ACT 3 — WHO'S SERVING IT  (supply landscape)

| # | Exhibit | Form | Source tab | The insight it must land | Conf | Deck |
|---|---------|------|-----------|--------------------------|------|------|
| 15 | Organised stores by tier | Bar | `11_CALC_Supply` | Supply clusters in top tiers | H | **S9** |
| 16 | Store mix by segment (Western/Ethnic/Dept/Sports/Footwear/Formal/Hyper) | Stacked bar | `6_IN_Stores` | Which formats exist where | H | **S10** |
| 17 | Store mix by ASP band (Economy→Luxury) | Stacked bar | `6_IN_Stores` | Price-point coverage | H | — |
| 18 | Segment × ASP store heatmap | Matrix heatmap | `6_IN_Stores` | Where supply piles up — and the empty cells | H | **S10** |
| 19 | Store density — stores per mn population, by tier | Bar | `11_CALC_Supply` | Lower tiers are structurally under-stored | H | **S12** |
| 20 | Listed-brand scorecard (stores, rev, rev/store, rev/sqft, area, adds) | Table | `5_IN_Brands` | Who the players are, at a glance | H/M | **S11** |
| 21 | Brand reach — unique pincodes / cities / states | Bar | `5_IN_Brands` | Footprint breadth vs depth | H | — |
| 22 | Store adds FY25 vs FY26 by brand | Paired bar | `5_IN_Brands` | Who's actually expanding (momentum) | H | — |
| 23 | Rev/store by tier × segment | Bar | `5_IN_Brands` + multipliers | The unit-econ gradient that prices the whitespace | M | — |
| 24 | Brands-per-pincode distribution | Histogram | supply | Competitive density — most pincodes are thin | H | — |

## ACT 4 — THE GAP  (whitespace = the thesis) ★ core act

| # | Exhibit | Form | Source tab | The insight it must land | Conf | Deck |
|---|---------|------|-----------|--------------------------|------|------|
| 25 | **WHITESPACE MATRIX by tier** (wallet \| captured \| penetration% \| whitespace \| implied headroom) | Matrix table | `OUT_Whitespace` | The whole thesis on one slide | M | **S14** |
| 26 | **Organised penetration S-curve by tier** | Line | `OUT_Whitespace` | Penetration collapses down-tier — that's the opening | M | **S13** |
| 27 | Whitespace ₹ by tier | Bar | `OUT_Whitespace` | Where the unserved rupees are | M | — |
| 28 | Implied store headroom by tier | Bar | `OUT_Whitespace` | Converts unserved ₹ into store count | M | **S14** |
| 29 | Whitespace by segment | Bar | `CALC` | Which formats have the most runway | M | **S15** |
| 30 | Penetration by segment × tier | Heatmap | `CALC` | The sharpest single gaps | M | — |
| 31 | Under-served clusters — store density vs per-capita wallet (bubble = wallet) | Scatter | `11_CALC_Supply` | The opportunity quadrant: high spend, low density | M | **S16** |
| 32 | **Value × Western whitespace by tier** (Zudio's lane) | Bar | `OUT_Whitespace` | Zoom from "market gap" to *Zudio's* gap | M | **S17** |
| 33 | Captured vs uncaptured wallet | 100% stacked | `OUT_Whitespace` | The headline penetration number | M | — |

## ACT 5 — PROOF  (comps validate the method)

| # | Exhibit | Form | Source tab | The insight it must land | Conf | Deck |
|---|---------|------|-----------|--------------------------|------|------|
| 34 | Comp table — listed peers (rev, stores, rev/store, rev/sqft, SSSG, adds) | Table | `7_IN_Comps` | Credibility anchor: the model's inputs match reality | H/M | **S18** |
| 35 | Rev/store vs peers | Bar | `7_IN_Comps` | Validates the rev/store assumption | M | **S19** |
| 36 | Rev/sqft vs peers | Bar | `7_IN_Comps` | Productivity benchmark | M | **S19** |
| 37 | FY26 store add-pace by peer | Bar | `7_IN_Comps` | What "sustainable pace" looks like in-market | H | — |
| 38 | Baazar Style DRHP benchmark (their stated TAM / penetration / runway) | Annotated table | `7_IN_Comps` | A third party reached the same conclusion independently | H | **S18** |
| 39 | Avg store size by peer | Bar | `7_IN_Comps` | Format calibration for the headroom math | M | — |
| 40 | Revenue CAGR vs store CAGR (peers) | Scatter | `7_IN_Comps` | Growth quality: density-led vs footprint-led | M | — |

## ACT 6 — ZUDIO STORE POTENTIAL  (the answer)

| # | Exhibit | Form | Source tab | The insight it must land | Conf | Deck |
|---|---------|------|-----------|--------------------------|------|------|
| 41 | **Store potential — 3 outputs × 3 scenarios** (saturation / runway / add-pace, Bear-Base-Bull) | Table | `OUT_Zudio` | The number, with its range | M | **S20** |
| 42 | Saturation store-count fan | Bar (Bear/Base/Bull) | `OUT_Zudio` | The spread of the prize | M | **S20** |
| 43 | Incremental runway vs current footprint | Bridge / bar | `OUT_Zudio` | How much room is left from here | M | **S21** |
| 44 | Runway in years — at current vs implied pace | Bar / line | `OUT_Zudio` | The time dimension of the opportunity | M | **S21** |
| 45 | Saturation stores by tier | Bar | `CALC` | *Where* the next N stores go | M | — |
| 46 | Sensitivity tornado — saturation vs rev/store, target pen, V×W share | Tornado | `1_Scenario` | What actually moves the answer (pre-empts the attack) | M | **S20** |
| 47 | Current footprint vs whitespace | Map / bar | `CALC` | Visual of the gap | M | — |
| 48 | Capturable wallet → stores conversion walk | Waterfall | `OUT_Zudio` | The logic chain, end to end | M | — |

## CLOSE

| # | Exhibit | Form | Source tab | The insight it must land | Conf | Deck |
|---|---------|------|-----------|--------------------------|------|------|
| 49 | Summary scorecard — TAM / whitespace / V×W runway / store potential / years | One-page dashboard | All | The whole argument in five numbers | — | **S22** |
| 50 | The recommendation / so-what | Statement slide | — | The one line — drawn from the insights, not pre-decided | — | **S1+S22** |

---

## The 22 that survive (the spine)

Open with the **answer** (1–2), establish **credibility/method** (S2 methodology + S18 comps), walk the **argument** (prize → demand → supply → gap), then **land Zudio** (S20–21) and **close** (S22). The thesis-carrying exhibits are **25, 26, 32, 41** — if a slide doesn't help one of those four land, it's an appendix candidate.

**Cut rule for getting from 50 → 22:** an exhibit earns a slide only if (a) its takeaway is a link in the chain, and (b) no other exhibit makes the same point more sharply. Duplicative "nice to know" geography/mix charts (9, 12, 13, 17, 21, 24, 39, 40, 47) are the first to move to appendix.

## Appendix bucket (built, not shown — kept for Q&A defence)
2, 6, 7, 9, 12, 13, 14, 17, 21, 22, 23, 24, 27, 30, 33, 36–37, 39, 40, 45, 47, 48 — these are exactly what a senior reviewer asks for *after* the main pitch, so build them but hold them back.
