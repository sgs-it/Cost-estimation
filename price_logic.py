"""
price_logic.py
==============
Pricing Logic System — derived from analysis of ALL Excel estimation sheets.

AMC FILES ANALYSED:
  - AMC - Asteco - Oberoi Tower - Landscaping AMC        (Outdoor, multi-sheet)
  - AMC - Index Tower - Renewal of Indoor AMC            (Indoor, renewal)
  - AMC - Wasl - Karama Shopping Complex - Palm Tree AMC (Outdoor/Palm, multi-sheet)
  - AMC - Engie - Bayti 20,33,40 - Landscaping AMC       (Outdoor, multi-sheet)

PROJECT FILES ANALYSED (PTI/PTO/VOI/VOO):
  - PTI Indoor Plant projects  → Pot + Plant + UnderPlant + Soil model
  - PTO Outdoor Plant projects → COST + LABOR + SUPRV + PLANT column model
  - VOI / VOO Variation orders → same models as PTI/PTO respectively

PRICING RULES EXTRACTED FROM SHEETS:
  AMC Outdoor:
    Labour Direct, Indirect → FTE × monthly salary × 12
    Equipment, PPE, Consumables, Subcontractor, Other → detailed line items
    Markup: Margin 10% + Overhead 10% + Nego 5% (excluding Other from Nego)
    Year 1 = Year 2; Year 3 = Labour escalated 5%, rest same

  AMC Indoor:
    Monthly = Σ(qty × unit_rate)
    Net Monthly = Monthly - discount
    VAT 5% on net monthly
    Annual = Net Monthly × 12

  Project Outdoor (PTO/VOO):
    Unit Price = COST + LABOR + SUPRV + PLANT + OH 10%
    Extended = Unit Price × QTY
    Total → discount → +VAT 5%

  Project Indoor (PTI/VOI):
    Sub-total Pot = Pot cost per item
    Sub-total Plant = Main Plant + Under Plants + Soil & Hydrostone
    OH Pot = Pot × 20%, OH Plant = Plant × 20%
    Combined = Pot + Plant
    Profit = Combined × 20%, Nego = Combined × 10%
    Unit Rate = Pot + Plant + OH_Pot + OH_Plant + Profit + Nego + Labor + Delivery
    Extended = Unit Rate × QTY
    Total → discount → +VAT 5%

STANDARD UNIT RATES (from real data):
  Labour rates (monthly all-in):
    Gardener Labour:     AED 1,100 (Basic 500 + OT 156 + Food 444)
    Gardener Expert:     AED 1,351 (Basic 675 + OT 211 + Food 465)
    Team Leader:         AED 1,800 (Basic 900 + OT 281 + Food 619)
    Irrigation Tech:     AED 1,800
    Palm Specialist:     AED 1,800
    Working Foreman:     AED 4,000
    Site Superintendent: AED 3,750 (indirect)
    Admin:               AED 4,000 (indirect)
    Landscaping Engr:    AED 8,000 (indirect)
    HSE Engineer:        AED 8,000 (indirect)

  Equivalent Area Factors (m2 per unit):
    Palm  → 13.5 m2 equivalent
    Tree  → 6.0  m2 equivalent
    Shrub / GC / Lawn → 1.0 m2 each

  Manpower Factor: 1 FTE per 4,000 m2 equiv. area (0.00025 FTE/m2)
  FTE monthly direct cost: AED 3,017.50 average blended
  FTE monthly indirect:    AED 275.35  average blended

  Equipment (annual per m2 equiv): AED 3.30
  PPE (annual per FTE):            AED 604.00 (min AED 302)
  Consumables base (annual/m2):    AED 3.15
  Seasonal flower rotation cost:   AED 99/lot × 3 rotations/year
  Subcontractor (annual/m2):       AED 1.40
  Insurance baseline:              AED 177.50/year (Third-party + WC + PI)

  Markup (applied to sub-total, fixed across all 3 years):
    Margin:      10%
    Overhead:    10%
    Negotiation: 5% (on sub-total EXCLUDING Other/Insurance)
    Total markup: ~25%

  VAT: 5% (UAE standard)
  Year 3 Escalation: 5% on Labour Direct + Labour Indirect only

  Indoor AMC unit rates (from Index Tower):
    Free-stand potted plant:  AED 15/pot/month
    Table plant arrangement:  AED 5/pot/month
    Discount applied:         AED 50/month (example)

  Project Outdoor Unit Rates (typical, from real sheets):
    Hibiscus shrub:              COST 7.5, LABOR 0.5, SUPRV 0, PLANT 0.5
    Paspalum lawn (m2):          COST 13, LABOR 3, SUPRV 0.75, PLANT 1
    Plumeria tree (std):         COST 308, LABOR 15, SUPRV 3, PLANT 30
    Neem tree:                   COST 282.5, LABOR 20, SUPRV 3.8, PLANT 50
    Bougainvillea:               COST 7, LABOR 0.5, SUPRV 0.125, PLANT 0.5
    Terminalia (std):            COST 310, LABOR 150, SUPRV 37.5, PLANT 50
    Terminalia (large):          COST 1060, LABOR 150, SUPRV 37.5, PLANT 50
    Delonix regia (std):         COST 310, LABOR 150, SUPRV 37.5, PLANT 50
    Delonix regia (large):       COST 1060, LABOR 150, SUPRV 37.5, PLANT 50
    Ficus nitida:                COST 710, LABOR 150, SUPRV 37.5, PLANT 50
    Ficus alii:                  COST 710, LABOR 150, SUPRV 37.5, PLANT 50
    Jatropha shrub:              COST 25, LABOR 5, SUPRV 1.25, PLANT 5
    Scaevola (shrub):            COST 15, LABOR 3, SUPRV 0.75, PLANT 2.5
    Ixora (shrub):               COST 22, LABOR 3, SUPRV 0.75, PLANT 2.5
    Pandanus (shrub):            COST 20, LABOR 3, SUPRV 0.75, PLANT 2.5
    Sanseveria (shrub):          COST 15, LABOR 3, SUPRV 0.75, PLANT 2.5
    Palm tree removal+install:   lumpsum from ~1340 AED
    Pebbles installation (LM):   COST 6, LABOR 1, PLANT 1, OH 10%
    White pebbles (bag):         COST 66, PLANT 2.78, OH 10%
    Water tanker (per trip):     ~AED 350-500

  Project Indoor Unit Rates (typical, from real sheets):
    GRP Ceramic pot (small):     Pot 190, Main Plant 120-275, Soil 55-65
    GRP Ceramic pot (large):     Pot 330-480, Main Plant 190-285, Soil 66-140
    Moisture indicator:          included in soil column when present
    OH rate on pot:              20%
    OH rate on plant/soil:       20%
    Profit margin:               20% on combined
    Nego buffer:                 10% on combined
"""

# ============================================================
# CONSTANTS — extracted directly from Excel sheet analysis
# ============================================================

# Markup rates (consistent across all AMC files)
MARGIN_PCT      = 0.10
OVERHEAD_PCT    = 0.10
NEGO_PCT        = 0.05   # Applied to subtotal MINUS insurance/other
VAT_PCT         = 0.05
YEAR3_LABOUR_ESC = 0.05

# Labour rates (blended averages from Labour sheets)
FTE_DIRECT_MONTHLY   = 3017.50   # Weighted average direct staff AED/month
FTE_INDIRECT_MONTHLY = 275.35    # Weighted average indirect share AED/month

# Area-based cost factors (annual, per m² equivalent area)
EQUIPMENT_PER_M2     = 3.30
CONSUMABLES_PER_M2   = 3.15
SUBCONTRACTOR_PER_M2 = 1.40

# PPE
PPE_PER_FTE  = 604.00
PPE_MINIMUM  = 302.00

# Equivalent area multipliers
PALM_EQ_M2   = 13.5
TREE_EQ_M2   = 6.0
SHRUB_EQ_M2  = 1.0
GC_EQ_M2     = 1.0
LAWN_EQ_M2   = 1.0

# Manpower factor: FTE per m² equivalent area
FTE_PER_M2 = 0.00025

# Insurance baseline (annual)
INSURANCE_BASE = 177.50

# Seasonal flower cost (per rotation, per 100 flowers)
FLOWER_RATE_PER_100 = 99.00
FLOWER_ROTATIONS    = 3

# Indoor AMC unit rates (monthly per pot)
INDOOR_LARGE_POT_RATE  = 15.00   # Free-stand potted plant
INDOOR_TABLE_RATE      = 5.00    # Table plant arrangement
INDOOR_WALL_RATE       = 10.00   # Wall-mounted / hanging

# Project Indoor overhead & markup
PTI_OH_POT_PCT     = 0.20
PTI_OH_PLANT_PCT   = 0.20
PTI_PROFIT_PCT     = 0.20
PTI_NEGO_PCT       = 0.10

# Project Outdoor overhead
PTO_OH_PCT = 0.10   # 10% on COST+PLANT combined


# ============================================================
# PRICING ENGINES
# ============================================================

class AmcOutdoorPricer:
    """
    Landscaping AMC pricing for outdoor/garden sites.
    Inputs: site physical asset inventory (palms, trees, shrubs, GC, lawn, flowers)
    Output: 3-year cost breakdown with markup, monthly rate, VAT
    """

    def __init__(self, config: dict):
        self.cfg = config

    def _equiv_area(self) -> float:
        c = self.cfg
        return (
            c.get("palms", 0) * PALM_EQ_M2 +
            c.get("trees", 0) * TREE_EQ_M2 +
            c.get("shrubs_m2", 0) * SHRUB_EQ_M2 +
            c.get("gc_m2", 0) * GC_EQ_M2 +
            c.get("lawn_m2", 0) * LAWN_EQ_M2
        )

    def _fte(self, eq_area: float) -> float:
        return eq_area * FTE_PER_M2

    def compute(self) -> dict:
        c = self.cfg
        eq_area = self._equiv_area()
        fte = self._fte(eq_area)

        flowers = c.get("seasonal_flowers", 0)
        flower_cost_annual = (flowers / 100) * FLOWER_RATE_PER_100 * FLOWER_ROTATIONS

        # --- Year 1 cost components ---
        labour_direct_y1   = fte * FTE_DIRECT_MONTHLY * 12
        labour_indirect_y1 = fte * FTE_INDIRECT_MONTHLY * 12
        equipment_y1       = eq_area * EQUIPMENT_PER_M2
        ppe_y1             = max(fte * PPE_PER_FTE, PPE_MINIMUM)
        consumables_y1     = (eq_area * CONSUMABLES_PER_M2) + flower_cost_annual
        subcontractor_y1   = c.get("subcontractor_annual", eq_area * SUBCONTRACTOR_PER_M2)
        admin_y1           = c.get("admin_annual", 0.0)
        other_y1           = c.get("insurance_annual", INSURANCE_BASE)

        subtotal_y1 = (labour_direct_y1 + labour_indirect_y1 + equipment_y1 +
                       ppe_y1 + consumables_y1 + subcontractor_y1 + admin_y1 + other_y1)

        # --- Year 2 = Year 1 (same costs) ---
        subtotal_y2 = subtotal_y1

        # --- Year 3 = Labour escalated 5%, rest same ---
        esc = 1 + YEAR3_LABOUR_ESC
        labour_direct_y3   = labour_direct_y1 * esc
        labour_indirect_y3 = labour_indirect_y1 * esc
        subtotal_y3 = (labour_direct_y3 + labour_indirect_y3 +
                       equipment_y1 + ppe_y1 + consumables_y1 +
                       subcontractor_y1 + admin_y1 + other_y1)

        # --- Markup (fixed, based on Y1, applied to all 3 years) ---
        margin_val   = subtotal_y1 * c.get("margin_pct", MARGIN_PCT)
        overhead_val = subtotal_y1 * c.get("overhead_pct", OVERHEAD_PCT)
        nego_base    = subtotal_y1 - other_y1   # Nego excludes insurance
        nego_val     = nego_base * c.get("nego_pct", NEGO_PCT)
        total_markup = margin_val + overhead_val + nego_val

        grand_y1 = subtotal_y1 + total_markup
        grand_y2 = subtotal_y2 + total_markup
        grand_y3 = subtotal_y3 + total_markup
        net_3yr  = grand_y1 + grand_y2 + grand_y3

        monthly_y1 = grand_y1 / 12
        vat_monthly = monthly_y1 * VAT_PCT

        return {
            # Site metrics
            "equiv_area_m2":       round(eq_area, 2),
            "fte_required":        round(fte, 4),
            # Year 1 breakdown
            "labour_direct_y1":    round(labour_direct_y1, 2),
            "labour_indirect_y1":  round(labour_indirect_y1, 2),
            "equipment_y1":        round(equipment_y1, 2),
            "ppe_y1":              round(ppe_y1, 2),
            "consumables_y1":      round(consumables_y1, 2),
            "subcontractor_y1":    round(subcontractor_y1, 2),
            "admin_y1":            round(admin_y1, 2),
            "other_y1":            round(other_y1, 2),
            "subtotal_y1":         round(subtotal_y1, 2),
            "subtotal_y2":         round(subtotal_y2, 2),
            "subtotal_y3":         round(subtotal_y3, 2),
            # Markup
            "margin_val":          round(margin_val, 2),
            "overhead_val":        round(overhead_val, 2),
            "nego_val":            round(nego_val, 2),
            "total_markup":        round(total_markup, 2),
            # Grand totals
            "grand_total_y1":      round(grand_y1, 2),
            "grand_total_y2":      round(grand_y2, 2),
            "grand_total_y3":      round(grand_y3, 2),
            "net_3yr_contract":    round(net_3yr, 2),
            # Monthly billing
            "monthly_rate_y1":     round(monthly_y1, 2),
            "vat_monthly":         round(vat_monthly, 2),
            "monthly_incl_vat":    round(monthly_y1 + vat_monthly, 2),
        }


class AmcIndoorPricer:
    """
    Indoor plant AMC pricing.
    Inputs: list of inventory items [{type, qty, unit_rate}], discount
    Output: monthly/annual charges + VAT
    """

    def __init__(self, config: dict):
        self.cfg = config

    def compute(self) -> dict:
        c = self.cfg
        items = c.get("items", [])

        # Build line-by-line subtotal
        base_monthly = 0.0
        line_details = []
        for item in items:
            qty       = item.get("qty", 0)
            # Use provided rate or fall back to type-based defaults
            rate = item.get("unit_rate", None)
            if rate is None:
                itype = item.get("type", "large").lower()
                if "table" in itype:
                    rate = INDOOR_TABLE_RATE
                elif "wall" in itype or "hanging" in itype:
                    rate = INDOOR_WALL_RATE
                else:
                    rate = INDOOR_LARGE_POT_RATE
            subtotal = qty * rate
            base_monthly += subtotal
            line_details.append({
                "description": item.get("description", "Plant Arrangement"),
                "qty": qty,
                "unit_rate": rate,
                "monthly_charge": round(subtotal, 2),
                "annual_charge": round(subtotal * 12, 2),
            })

        discount      = c.get("discount_monthly", 0.0)
        net_monthly   = base_monthly - discount
        net_annual    = net_monthly * 12
        vat_monthly   = net_monthly * VAT_PCT
        gross_monthly = net_monthly + vat_monthly

        return {
            "line_details":       line_details,
            "base_monthly":       round(base_monthly, 2),
            "base_annual":        round(base_monthly * 12, 2),
            "discount_monthly":   round(discount, 2),
            "net_monthly":        round(net_monthly, 2),
            "net_annual":         round(net_annual, 2),
            "vat_monthly":        round(vat_monthly, 2),
            "gross_monthly_incl_vat": round(gross_monthly, 2),
        }


class ProjectOutdoorPricer:
    """
    Project / Variation Order pricing for OUTDOOR works (PTO / VOO).
    Inputs: list of line items [{description, qty, unit, cost, labor, suprv, plant}]
    Output: itemised quote with OH 10%, totals, VAT
    """

    def __init__(self, config: dict):
        self.cfg = config

    def compute(self) -> dict:
        c = self.cfg
        items = c.get("items", [])
        oh_pct = c.get("oh_pct", PTO_OH_PCT)

        manifest = []
        gross_total = 0.0

        for item in items:
            qty    = item.get("qty", 1)
            cost   = item.get("cost", 0.0)
            labor  = item.get("labor", 0.0)
            suprv  = item.get("suprv", 0.0)
            plant  = item.get("plant", 0.0)

            material = cost + plant                    # OH applied on material+plant
            oh       = material * oh_pct
            unit_rate = round(cost + labor + suprv + plant + oh)
            extended  = unit_rate * qty
            gross_total += extended

            manifest.append({
                "description": item.get("description", "Works"),
                "qty":         qty,
                "unit":        item.get("unit", "nos"),
                "cost":        cost,
                "labor":       labor,
                "suprv":       suprv,
                "plant":       plant,
                "oh":          round(oh, 2),
                "unit_rate":   unit_rate,
                "extended":    extended,
            })

        discount    = c.get("discount", 0.0)
        taxable     = gross_total - discount
        vat         = taxable * VAT_PCT
        final_quote = taxable + vat

        return {
            "manifest":      manifest,
            "gross_total":   round(gross_total, 2),
            "discount":      round(discount, 2),
            "taxable_base":  round(taxable, 2),
            "vat":           round(vat, 2),
            "final_quote":   round(final_quote, 2),
        }


class ProjectIndoorPricer:
    """
    Project / Variation Order pricing for INDOOR works (PTI / VOI).
    Inputs: list of line items [{description, qty, pot, main_plant, under_plants,
                                  soil_hydrostone, labor, delivery}]
    Output: itemised quote with full markup structure (OH 20%+20%, Profit 20%, Nego 10%)
    """

    def __init__(self, config: dict):
        self.cfg = config

    def compute(self) -> dict:
        c = self.cfg
        items = c.get("items", [])
        oh_pot_pct   = c.get("oh_pot_pct",   PTI_OH_POT_PCT)
        oh_plant_pct = c.get("oh_plant_pct", PTI_OH_PLANT_PCT)
        profit_pct   = c.get("profit_pct",   PTI_PROFIT_PCT)
        nego_pct     = c.get("nego_pct",     PTI_NEGO_PCT)

        manifest = []
        gross_total = 0.0

        for item in items:
            qty             = item.get("qty", 1)
            pot             = item.get("pot", 0.0)
            main_plant      = item.get("main_plant", 0.0)
            under_plants    = item.get("under_plants", 0.0)
            soil_hydrostone = item.get("soil_hydrostone", 0.0)
            labor           = item.get("labor", 0.0)
            delivery        = item.get("delivery", 0.0)

            sub_pot   = pot
            sub_plant = main_plant + under_plants + soil_hydrostone
            combined  = sub_pot + sub_plant

            oh_pot    = sub_pot   * oh_pot_pct
            oh_plant  = sub_plant * oh_plant_pct
            profit    = combined  * profit_pct
            nego      = combined  * nego_pct

            unit_cost = combined + oh_pot + oh_plant + profit + nego + labor + delivery
            unit_rate = round(unit_cost)       # Integer rounding per sheet convention
            extended  = unit_rate * qty
            gross_total += extended

            manifest.append({
                "description":   item.get("description", "Indoor Plant Supply & Install"),
                "qty":           qty,
                "pot":           pot,
                "main_plant":    main_plant,
                "under_plants":  under_plants,
                "soil_hydrostone": soil_hydrostone,
                "labor":         labor,
                "delivery":      delivery,
                "oh_pot":        round(oh_pot, 2),
                "oh_plant":      round(oh_plant, 2),
                "profit":        round(profit, 2),
                "nego":          round(nego, 2),
                "unit_cost_raw": round(unit_cost, 2),
                "unit_rate":     unit_rate,
                "extended":      extended,
            })

        discount    = c.get("discount", 0.0)
        taxable     = gross_total - discount
        vat         = taxable * VAT_PCT
        final_quote = taxable + vat

        return {
            "manifest":      manifest,
            "gross_total":   round(gross_total, 2),
            "discount":      round(discount, 2),
            "taxable_base":  round(taxable, 2),
            "vat":           round(vat, 2),
            "final_quote":   round(final_quote, 2),
        }


# ============================================================
# DISPATCHER — choose engine by job type
# ============================================================

def get_estimate(job_type: str, config: dict) -> dict:
    """
    job_type options:
      "amc_outdoor"  → AmcOutdoorPricer
      "amc_indoor"   → AmcIndoorPricer
      "project_outdoor" → ProjectOutdoorPricer
      "project_indoor"  → ProjectIndoorPricer
    """
    engines = {
        "amc_outdoor":       AmcOutdoorPricer,
        "amc_indoor":        AmcIndoorPricer,
        "project_outdoor":   ProjectOutdoorPricer,
        "project_indoor":    ProjectIndoorPricer,
    }
    engine_cls = engines.get(job_type)
    if engine_cls is None:
        raise ValueError(f"Unknown job_type '{job_type}'. Choose from: {list(engines.keys())}")
    return engine_cls(config).compute()


# ============================================================
# QUICK DEMO (run directly to verify)
# ============================================================
if __name__ == "__main__":
    print("\n=== AMC OUTDOOR DEMO (Oberoi Tower equivalent) ===")
    result = get_estimate("amc_outdoor", {
        "palms": 6, "trees": 8, "shrubs_m2": 200, "gc_m2": 100,
        "lawn_m2": 0, "seasonal_flowers": 100,
    })
    for k, v in result.items():
        print(f"  {k:<30}: {v:,}" if isinstance(v, (int,float)) else f"  {k:<30}: {v}")

    print("\n=== AMC INDOOR DEMO (Index Tower equivalent) ===")
    result = get_estimate("amc_indoor", {
        "items": [
            {"description": "LG Level - Free Stand Potted Plants", "qty": 2, "unit_rate": 15},
            {"description": "Ground Level - Free Stand Potted Plants", "qty": 14, "unit_rate": 15},
            {"description": "29th Level - Free Stand Potted Plants", "qty": 12, "unit_rate": 15},
            {"description": "29th Level - Table Plant Arrangements", "qty": 6, "unit_rate": 5},
        ],
        "discount_monthly": 50,
    })
    for k, v in result.items():
        if k != "line_details":
            print(f"  {k:<35}: {v:,}" if isinstance(v, (int,float)) else f"  {k:<35}: {v}")

    print("\n=== PROJECT OUTDOOR DEMO (Vandalism Replacement) ===")
    result = get_estimate("project_outdoor", {
        "items": [
            {"description": "Supply & Install Hibiscus", "qty": 120, "unit": "nos",
             "cost": 7.5, "labor": 0.5, "suprv": 0, "plant": 0.5},
            {"description": "Supply & Install Paspalum Lawn", "qty": 10, "unit": "m2",
             "cost": 13, "labor": 3, "suprv": 0.75, "plant": 1},
            {"description": "Supply & Install Plumeria Tree", "qty": 1, "unit": "nos",
             "cost": 308, "labor": 15, "suprv": 3, "plant": 30},
        ],
        "discount": 0,
    })
    for k, v in result.items():
        if k != "manifest":
            print(f"  {k:<30}: {v:,}" if isinstance(v, (int,float)) else f"  {k:<30}: {v}")

    print("\n=== PROJECT INDOOR DEMO (Ministry of Sports) ===")
    result = get_estimate("project_indoor", {
        "items": [
            {"description": "Supply & Install GRP Pot + Ficus", "qty": 1,
             "pot": 190, "main_plant": 120, "under_plants": 0, "soil_hydrostone": 55,
             "labor": 0, "delivery": 0},
        ],
        "discount": 0,
    })
    for k, v in result.items():
        if k != "manifest":
            print(f"  {k:<30}: {v:,}" if isinstance(v, (int,float)) else f"  {k:<30}: {v}")
