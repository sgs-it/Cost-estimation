import os
import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import openpyxl
from price_logic import get_estimate
from excel_exporter import generate_excel

def test_amc_outdoor():
    print("Testing AMC Outdoor...")
    config = {
        "palms": 10,
        "trees": 20,
        "shrubs_m2": 500,
        "gc_m2": 200,
        "lawn_m2": 300,
        "seasonal_flowers": 150,
        "margin_pct": 0.10,
        "overhead_pct": 0.10,
        "nego_pct": 0.05
    }
    # Test without total_space
    res1 = get_estimate("amc_outdoor", config)
    print(f"  FTE Required: {res1['fte_required']}")
    print(f"  Equiv Area: {res1['equiv_area_m2']}")
    print(f"  Year 1 Grand Total: {res1['grand_total_y1']}")
    print(f"  Total Space (Default): {res1['total_space']}")

    # Test with total_space override
    config["total_space"] = 2500
    res2 = get_estimate("amc_outdoor", config)
    print(f"  With total_space override (2500 m²):")
    print(f"    FTE Required (should be same as before): {res2['fte_required']}")
    print(f"    Equiv Area (should be same as before): {res2['equiv_area_m2']}")
    print(f"    Total Space: {res2['total_space']}")
    print(f"    Year 1 Grand Total (should change because consumables, eq, etc. use total_space): {res2['grand_total_y1']}")
    assert res1["fte_required"] == res2["fte_required"]
    assert res1["grand_total_y1"] != res2["grand_total_y1"]
    
    # Export to excel test
    meta = {"site": "Test Site", "client": "Test Client"}
    xlsx_bytes = generate_excel("amc_outdoor", res2, config, meta)
    assert len(xlsx_bytes) > 0
    print("  AMC Outdoor Excel exported successfully.")

def test_project_outdoor():
    print("Testing Project Outdoor...")
    config = {
        "items": [
            {"description": "Damas Tree Pruning", "qty": 5, "unit": "nos",
             "cost": 100, "labor": 50, "suprv": 12.5, "plant": 20},
            {"description": "Lawn Mowing", "qty": 100, "unit": "m2",
             "cost": 5, "labor": 2, "suprv": 0.5, "plant": 1}
        ],
        "oh_material_pct": 0.20,
        "oh_labor_pct": 0.20,
        "profit_material_pct": 0.20,
        "profit_labor_pct": 0.20,
        "nego_pct": 0.10,
        "suprv_pct": 0.25,
        "discount": 100
    }
    res = get_estimate("project_outdoor", config)
    print(f"  Gross Total: {res['gross_total']}")
    print(f"  Taxable Base: {res['taxable_base']}")
    print(f"  Final Quote: {res['final_quote']}")
    assert len(res["manifest"]) == 2
    
    # Check item details
    item1 = res["manifest"][0]
    print(f"  Item 1 markups: OH Mat={item1['oh_material']}, OH Lab={item1['oh_labor']}, Profit Mat={item1['profit_material']}, Profit Lab={item1['profit_labor']}, Nego={item1['nego']}")
    print(f"  Item 1 unit rate (rounded/ceiled): {item1['unit_rate']}")
    
    # Export to excel test
    meta = {"site": "Test Site", "client": "Test Client"}
    xlsx_bytes = generate_excel("project_outdoor", res, config, meta)
    assert len(xlsx_bytes) > 0
    print("  Project Outdoor Excel exported successfully.")

def test_project_indoor():
    print("Testing Project Indoor...")
    config = {
        "items": [
            {"description": "Ministry of Sports Entry Planter", "qty": 10,
             "pot": 190, "main_plant": 275, "under_plants": 0, "soil_hydrostone": 58,
             "labor": 30, "delivery": 20}
        ],
        "oh_pot_pct": 0.20,
        "oh_plant_pct": 0.20,
        "profit_pct": 0.20,
        "nego_pct": 0.10,
        "discount": 50
    }
    res = get_estimate("project_indoor", config)
    print(f"  Gross Total: {res['gross_total']}")
    print(f"  Taxable Base: {res['taxable_base']}")
    print(f"  Final Quote: {res['final_quote']}")
    assert len(res["manifest"]) == 1
    
    # Check item details
    item = res["manifest"][0]
    print(f"  Item markups: OH Pot={item['oh_pot']}, OH Plant={item['oh_plant']}, Profit={item['profit']}, Nego={item['nego']}")
    print(f"  Item unit rate (should keep decimals): {item['unit_rate']}")
    
    # Export to excel test
    meta = {"site": "Test Site", "client": "Test Client"}
    xlsx_bytes = generate_excel("project_indoor", res, config, meta)
    assert len(xlsx_bytes) > 0
    print("  Project Indoor Excel exported successfully.")

if __name__ == "__main__":
    test_amc_outdoor()
    test_project_outdoor()
    test_project_indoor()
    print("\nALL TESTS PASSED SUCCESSFULLY!")
