import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_big_sam_menu():
    wb = openpyxl.Workbook()
    
    # Remove default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)
    
    # Color Palette - Premium Navy & Teal Theme
    NAVY = "1B365D"
    TEAL = "008080"
    LIGHT_GRAY = "F2F4F7"
    ALT_ROW = "F9FAFB"
    WHITE = "FFFFFF"
    
    # Special Highlights
    ALLERGY_BG = "FFF2CC"    # Soft yellow
    ALLERGY_TXT = "800000"   # Dark red
    ADULT_BG = "FFE5CC"      # Soft orange
    ADULT_TXT = "B35A00"     # Dark orange
    TOTAL_BG = "E6EEF8"      # Soft blue-gray for totals
    
    # Fonts
    font_title = Font(name="Segoe UI", size=16, bold=True, color=NAVY)
    font_subtitle = Font(name="Segoe UI", size=11, italic=True, color="555555")
    font_header = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
    font_subheader = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
    font_bold = Font(name="Segoe UI", size=10, bold=True)
    font_regular = Font(name="Segoe UI", size=10)
    font_italic = Font(name="Segoe UI", size=10, italic=True)
    font_allergy = Font(name="Segoe UI", size=10, bold=True, color=ALLERGY_TXT)
    font_adult = Font(name="Segoe UI", size=10, bold=True, color=ADULT_TXT)
    
    # Fills
    fill_navy = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    fill_teal = PatternFill(start_color=TEAL, end_color=TEAL, fill_type="solid")
    fill_light_gray = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
    fill_alt_row = PatternFill(start_color=ALT_ROW, end_color=ALT_ROW, fill_type="solid")
    fill_allergy = PatternFill(start_color=ALLERGY_BG, end_color=ALLERGY_BG, fill_type="solid")
    fill_adult = PatternFill(start_color=ADULT_BG, end_color=ADULT_BG, fill_type="solid")
    fill_total = PatternFill(start_color=TOTAL_BG, end_color=TOTAL_BG, fill_type="solid")
    
    # Borders
    thin_side = Side(border_style="thin", color="D3D3D3")
    double_side = Side(border_style="double", color="333333")
    thick_top = Side(border_style="medium", color="1B365D")
    
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    # Alignments
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    align_wrap_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 1. MEAL PLAN DATA DEFINITION
    # Phase 1: Days 1-3 (Initial Carry)
    # Phase 2: Days 4-6 (Mule Drop)
    menu_data = [
        # --- PHASE 1 ---
        {
            "phase": 1,
            "day": 1,
            "date_str": "Sat Jul 25",
            "meals": {
                "BREAKFAST": [],  # Eaten before departure
                "SNACKS": [
                    ["Kind – Peanut Butter Dark Chocolate Bar", 3, 4, 3, 200, 1.04, 1.4, "⚠ NUT: swap → 88 Acres bar", True, False],
                    ["Kind – Caramel Almond & Peanut Bar", 3, 4, 3, 180, 1.59, 1.4, "⚠ NUT: swap → 88 Acres bar", True, False],
                    ["Krave Sea Salt Original Beef Jerky", 3, 4, 3, 140, 2.19, 1.48, "Safe for nut allergy ✓", False, False],
                    ["Gatorade Powder Pack", 3, 4, 3, 130, 0.38, 1.23, "Electrolytes", False, False],
                    ["Trader Joe's Dried Mango Slices (1 oz)", 3, 4, 3, 90, 0.55, 1.0, "TJ's — Sat snack · nut-free ✓", False, False]
                ],
                "LUNCH": [
                    ["ProBar Meal – Oatmeal Chocolate Chip", 3, 4, 3, 290, 2.99, 3.0, "Lunch · ⚠ NUT (cashews)", True, False]
                ],
                "DINNER": [
                    ["Peak Refuel – Beef Pasta Marinara", 2, 3, 3, 780, 7.49, 3.2, "Sat dinner (1.5 svgs/pouch)", False, False],
                    ["Mission Street Taco Flour Tortillas", 6, 8, 6, 75, 0.25, 0.5, "2/person; nut-free ✓", False, False],
                    ["Tillamook Cheddar Cheese Packets", 6, 8, 6, 90, 0.75, 0.7, "2/person; nut-free ✓", False, False],
                    ["Swiss Miss Hot Chocolate", 6, 8, 0, 90, 0.38, 0.73, "Scouts only (2 packets/person)", False, False]
                ],
                "DESSERT": [
                    ["Backpacker's Pantry – Crème Brûlée", 2, 2, 2, 620, 8.99, 2.26, "Buddy-pair shares pouch", False, False],
                    ["Swiss Miss Hot Chocolate", 6, 8, 0, 90, 0.38, 0.73, "Scouts only (2 packets/person)", False, False]
                ]
            }
        },
        {
            "phase": 1,
            "day": 2,
            "date_str": "Sun Jul 26",
            "meals": {
                "BREAKFAST": [
                    ["MET-Rx Big100 – Super Cookie Crunch", 3, 4, 3, 410, 2.19, 3.53, "Bkfst · ⚠ NUT", True, False],
                    ["Swiss Miss Hot Chocolate", 6, 8, 0, 90, 0.38, 0.73, "Scouts only (2 packets/person)", False, False],
                    ["Adults: own coffee / tea", 0, 0, 3, 0, 0.0, 0.0, "Adults only ★", False, True]
                ],
                "SNACKS": [
                    ["Krave Sea Salt Original Beef Jerky", 3, 4, 3, 140, 2.19, 1.48, "Safe for nut allergy ✓", False, False],
                    ["Kind – Dark Choc Nuts & Sea Salt Bar", 3, 4, 3, 200, 1.04, 1.4, "⚠ NUT: swap → 88 Acres bar", True, False],
                    ["Gatorade Powder Pack", 6, 8, 6, 130, 0.38, 1.23, "×2/person; climb day", False, False],
                    ["Banana Chips (1 oz)", 3, 4, 3, 100, 0.45, 1.0, "Nut-free ✓", False, False]
                ],
                "LUNCH": [
                    ["Chicken of the Sea Shredded Chicken Pouch", 3, 4, 3, 70, 2.29, 2.5, "No cook; nut-free ✓", False, False],
                    ["Mission Street Taco Flour Tortillas", 6, 8, 6, 75, 0.25, 0.5, "2/person; nut-free ✓", False, False],
                    ["Tillamook Cheddar Cheese Packets", 3, 4, 3, 90, 0.75, 0.7, "1/person; nut-free ✓", False, False],
                    ["Clif Bar – Crunchy PB", 3, 4, 3, 250, 1.29, 2.4, "Lunch snack; ⚠ NUT", True, False],
                    ["Justin's Classic PB Squeeze Pack", 3, 4, 3, 190, 1.49, 1.15, "Lunch side; ⚠ NUT", True, False]
                ],
                "SUPPLEMENT": [
                    ["Clif Builder's Bar (adult supplement)", 0, 0, 3, 290, 2.17, 2.4, "Adults only ★; ⚠ NUT", True, True],
                    ["Justin's PB Squeeze (adult supplement)", 0, 0, 3, 190, 1.49, 1.15, "Adults only ★; ⚠ NUT", True, True]
                ],
                "DINNER": [
                    ["Peak Refuel – Beef Stroganoff", 2, 3, 3, 600, 7.49, 3.5, "Sun dinner (1.5 svgs/pouch)", False, False],
                    ["Mission Street Taco Flour Tortillas", 6, 8, 6, 75, 0.25, 0.5, "2/person; nut-free ✓", False, False],
                    ["Idahoan Mashed Potatoes – Butter & Herb", 3, 4, 3, 110, 0.98, 1.0, "Dinner side", False, False],
                    ["Swiss Miss Hot Chocolate", 6, 8, 0, 90, 0.38, 0.73, "Scouts only (2 packets/person)", False, False]
                ],
                "DESSERT": [
                    ["Justin's Dark Chocolate Peanut Butter Cups", 3, 4, 3, 210, 1.49, 1.41, "1 packet/person; ⚠ NUT", True, False],
                    ["Swiss Miss Hot Chocolate", 6, 8, 0, 90, 0.38, 0.73, "Scouts only (2 packets/person)", False, False]
                ]
            }
        },
        {
            "phase": 1,
            "day": 3,
            "date_str": "Mon Jul 27",
            "meals": {
                "BREAKFAST": [
                    ["ProBar Meal – Oatmeal Chocolate Chip", 3, 4, 3, 290, 2.99, 3.0, "Bkfst · ⚠ NUT (cashews)", True, False],
                    ["Swiss Miss Hot Chocolate", 6, 8, 0, 90, 0.38, 0.73, "Scouts only (2 packets/person)", False, False],
                    ["Adults: own coffee / tea", 0, 0, 3, 0, 0.0, 0.0, "Adults only ★", False, True]
                ],
                "SNACKS": [
                    ["Krave Sea Salt Original Beef Jerky", 3, 4, 3, 140, 2.19, 1.48, "Safe for nut allergy ✓", False, False],
                    ["Clif Bar – Crunchy PB", 3, 4, 3, 250, 1.29, 2.4, "⚠ NUT", True, False],
                    ["Gatorade Powder Pack", 3, 4, 3, 130, 0.38, 1.23, "Electrolytes", False, False],
                    ["Peak Refuel – Fudge Brownie Bites", 2, 2, 2, 450, 4.99, 1.5, "Buddy-pair shares pouch", False, False]
                ],
                "LUNCH": [
                    ["Columbus Hard Salami (sliced packet)", 3, 4, 3, 110, 2.99, 2.0, "Lunch; 1/person; nut-free", False, False],
                    ["Wasa Crisp'n Light Crackerbread", 3, 4, 3, 60, 0.50, 1.0, "Lunch side", False, False],
                    ["Mission Street Taco Flour Tortillas", 6, 8, 6, 75, 0.25, 0.5, "2/person; nut-free ✓", False, False],
                    ["Tillamook Cheddar Cheese Packets", 3, 4, 3, 90, 0.75, 0.7, "1/person; nut-free ✓", False, False],
                    ["Justin's Classic PB Squeeze Pack", 3, 4, 3, 190, 1.49, 1.15, "Lunch side; ⚠ NUT", True, False]
                ],
                "DINNER": [
                    ["Peak Refuel – Beef Pasta Marinara", 2, 3, 3, 780, 7.49, 3.2, "Mon dinner (1.5 svgs/pouch)", False, False],
                    ["Mission Street Taco Flour Tortillas", 6, 8, 6, 75, 0.25, 0.5, "2/person; nut-free ✓", False, False],
                    ["Tillamook Cheddar Cheese Packets", 6, 8, 6, 90, 0.75, 0.7, "2/person; nut-free ✓", False, False],
                    ["Swiss Miss Hot Chocolate", 6, 8, 0, 90, 0.38, 0.73, "Scouts only (2 packets/person)", False, False]
                ],
                "DESSERT": [
                    ["Backpacker's Pantry – Crème Brûlée", 2, 2, 2, 620, 8.99, 2.26, "Buddy-pair shares pouch", False, False],
                    ["Swiss Miss Hot Chocolate", 6, 8, 0, 90, 0.38, 0.73, "Scouts only (2 packets/person)", False, False]
                ]
            }
        },
        # --- PHASE 2 ---
        {
            "phase": 2,
            "day": 4,
            "date_str": "Tue Jul 28",
            "meals": {
                "BREAKFAST": [
                    ["MET-Rx Big100 – Super Cookie Crunch", 3, 4, 3, 410, 2.19, 3.53, "Bkfst · ⚠ NUT", True, False],
                    ["Swiss Miss Hot Chocolate", 6, 8, 0, 90, 0.38, 0.73, "Scouts only (2 packets/person)", False, False],
                    ["Adults: own coffee / tea", 0, 0, 3, 0, 0.0, 0.0, "Adults only ★", False, True]
                ],
                "SNACKS": [
                    ["Krave Sea Salt Original Beef Jerky", 3, 4, 3, 140, 2.19, 1.48, "Safe for nut allergy ✓", False, False],
                    ["Kind – Dark Choc Nuts & Sea Salt Bar", 3, 4, 3, 200, 1.04, 1.4, "⚠ NUT: swap → 88 Acres bar", True, False],
                    ["Gatorade Powder Pack", 6, 8, 6, 130, 0.38, 1.23, "×2/person; climb day", False, False],
                    ["Banana Chips (1 oz)", 3, 4, 3, 100, 0.45, 1.0, "Nut-free ✓", False, False]
                ],
                "LUNCH": [
                    ["ProBar Meal – Oatmeal Chocolate Chip", 3, 4, 3, 290, 2.99, 3.0, "Lunch · ⚠ NUT (cashews)", True, False]
                ],
                "DINNER": [
                    ["Peak Refuel – Beef Stroganoff", 2, 3, 3, 600, 7.49, 3.5, "Tue dinner (1.5 svgs/pouch)", False, False],
                    ["Mission Street Taco Flour Tortillas", 6, 8, 6, 75, 0.25, 0.5, "2/person; nut-free ✓", False, False],
                    ["Idahoan Mashed Potatoes – Butter & Herb", 3, 4, 3, 110, 0.98, 1.0, "Dinner side", False, False],
                    ["Swiss Miss Hot Chocolate", 6, 8, 0, 90, 0.38, 0.73, "Scouts only (2 packets/person)", False, False]
                ],
                "DESSERT": [
                    ["Justin's Dark Chocolate Peanut Butter Cups", 3, 4, 3, 210, 1.49, 1.41, "1 packet/person; ⚠ NUT", True, False],
                    ["Swiss Miss Hot Chocolate", 6, 8, 0, 90, 0.38, 0.73, "Scouts only (2 packets/person)", False, False]
                ]
            }
        },
        {
            "phase": 2,
            "day": 5,
            "date_str": "Wed Jul 29",
            "meals": {
                "BREAKFAST": [
                    ["ProBar Meal – Oatmeal Chocolate Chip", 3, 4, 3, 290, 2.99, 3.0, "Bkfst · ⚠ NUT (cashews)", True, False],
                    ["Swiss Miss Hot Chocolate", 6, 8, 0, 90, 0.38, 0.73, "Scouts only (2 packets/person)", False, False],
                    ["Adults: own coffee / tea", 0, 0, 3, 0, 0.0, 0.0, "Adults only ★", False, True]
                ],
                "SNACKS": [
                    ["Kind – Peanut Butter Dark Chocolate Bar", 3, 4, 3, 200, 1.04, 1.4, "⚠ NUT: swap → 88 Acres bar", True, False],
                    ["Kind – Caramel Almond & Peanut Bar", 3, 4, 3, 180, 1.59, 1.4, "⚠ NUT: swap → 88 Acres bar", True, False],
                    ["Krave Sea Salt Original Beef Jerky", 3, 4, 3, 140, 2.19, 1.48, "Safe for nut allergy ✓", False, False],
                    ["Gatorade Powder Pack", 3, 4, 3, 130, 0.38, 1.23, "Electrolytes", False, False],
                    ["Trader Joe's Dried Mango Slices (1 oz)", 3, 4, 3, 90, 0.55, 1.0, "TJ's — Wed snack · nut-free ✓", False, False]
                ],
                "LUNCH": [
                    ["Chicken of the Sea Shredded Chicken Pouch", 3, 4, 3, 70, 2.29, 2.5, "No cook; nut-free ✓", False, False],
                    ["Mission Street Taco Flour Tortillas", 6, 8, 6, 75, 0.25, 0.5, "2/person; nut-free ✓", False, False],
                    ["Tillamook Cheddar Cheese Packets", 3, 4, 3, 90, 0.75, 0.7, "1/person; nut-free ✓", False, False],
                    ["Clif Bar – Crunchy PB", 3, 4, 3, 250, 1.29, 2.4, "Lunch snack; ⚠ NUT", True, False],
                    ["Justin's Classic PB Squeeze Pack", 3, 4, 3, 190, 1.49, 1.15, "Lunch side; ⚠ NUT", True, False]
                ],
                "SUPPLEMENT": [
                    ["Clif Builder's Bar (adult supplement)", 0, 0, 3, 290, 2.17, 2.4, "Adults only ★; ⚠ NUT", True, True],
                    ["Justin's PB Squeeze (adult supplement)", 0, 0, 3, 190, 1.49, 1.15, "Adults only ★; ⚠ NUT", True, True]
                ],
                "DINNER": [
                    ["Peak Refuel – Beef Pasta Marinara", 2, 3, 3, 780, 7.49, 3.2, "Wed dinner (1.5 svgs/pouch)", False, False],
                    ["Mission Street Taco Flour Tortillas", 6, 8, 6, 75, 0.25, 0.5, "2/person; nut-free ✓", False, False],
                    ["Tillamook Cheddar Cheese Packets", 6, 8, 6, 90, 0.75, 0.7, "2/person; nut-free ✓", False, False],
                    ["Swiss Miss Hot Chocolate", 6, 8, 0, 90, 0.38, 0.73, "Scouts only (2 packets/person)", False, False]
                ],
                "DESSERT": [
                    ["Backpacker's Pantry – Crème Brûlée", 2, 2, 2, 620, 8.99, 2.26, "Buddy-pair shares pouch", False, False],
                    ["Swiss Miss Hot Chocolate", 6, 8, 0, 90, 0.38, 0.73, "Scouts only (2 packets/person)", False, False]
                ]
            }
        },
        {
            "phase": 2,
            "day": 6,
            "date_str": "Thu Jul 30",
            "meals": {
                "BREAKFAST": [
                    ["MET-Rx Big100 – Super Cookie Crunch", 3, 4, 3, 410, 2.19, 3.53, "Bkfst · ⚠ NUT", True, False],
                    ["Swiss Miss Hot Chocolate", 6, 8, 0, 90, 0.38, 0.73, "Scouts only (2 packets/person)", False, False],
                    ["Adults: own coffee / tea", 0, 0, 3, 0, 0.0, 0.0, "Adults only ★", False, True]
                ],
                "SNACKS": [
                    ["Krave Sea Salt Original Beef Jerky", 3, 4, 3, 140, 2.19, 1.48, "Safe for nut allergy ✓", False, False],
                    ["Clif Bar – Crunchy PB", 3, 4, 3, 250, 1.29, 2.4, "⚠ NUT", True, False],
                    ["Gatorade Powder Pack", 3, 4, 3, 130, 0.38, 1.23, "Electrolytes", False, False],
                    ["Peak Refuel – Fudge Brownie Bites", 2, 2, 2, 450, 4.99, 1.5, "Buddy-pair shares pouch", False, False]
                ],
                "LUNCH": [
                    ["Columbus Hard Salami (sliced packet)", 3, 4, 3, 110, 2.99, 2.0, "Lunch; 1/person; nut-free", False, False],
                    ["Wasa Crisp'n Light Crackerbread", 3, 4, 3, 60, 0.50, 1.0, "Lunch side", False, False],
                    ["Mission Street Taco Flour Tortillas", 6, 8, 6, 75, 0.25, 0.5, "2/person; nut-free ✓", False, False],
                    ["Tillamook Cheddar Cheese Packets", 3, 4, 3, 90, 0.75, 0.7, "1/person; nut-free ✓", False, False],
                    ["Justin's Classic PB Squeeze Pack", 3, 4, 3, 190, 1.49, 1.15, "Lunch side; ⚠ NUT", True, False]
                ],
                "DINNER": [],
                "DESSERT": []
            }
        }
    ]

    # Initialize Sheet Row Pointers & Excel references
    day_total_rows = {}
    phase_total_rows = {}
    grand_total_row = None
    shopping_list_items_rows = {}

    # --- 2. MEAL PLAN (PER BEAR CAN) SHEET ---
    ws_mp = wb.create_sheet(title="Meal Plan (Per Bear Can)")
    ws_mp.views.sheetView[0].showGridLines = True
    
    ws_mp["A1"] = "Meal Plan — Quantities & Nutritional Breakdown Per Bear Can Type"
    ws_mp["A1"].font = font_title
    ws_mp["A2"] = "Calculated for: Scout Can (3 people), Scout Can (4 people), and Adult Can (3 people)  |  ⚠ Highlighted = Contains Nuts (swap for allergy scout)"
    ws_mp["A2"].font = font_subtitle
    
    headers = [
        "Food Item", "Qty\n(Scout 3)", "Qty\n(Scout 4)", "Qty\n(Adult 3)",
        "kcal/unit", "Scout 3\nkcal", "Scout 4\nkcal", "Adult 3\nkcal",
        "$/unit", "Scout 3\n$", "Scout 4\n$", "Adult 3\n$",
        "oz/unit", "Scout 3\noz", "Scout 4\noz", "Adult 3\noz", "Notes"
    ]
    
    def write_subheader(ws, r, title_text, fill_color):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=17)
        cell = ws.cell(row=r, column=1, value=title_text)
        cell.font = font_subheader
        cell.fill = fill_color
        cell.alignment = align_left
        ws.row_dimensions[r].height = 24
        
    def write_table_headers(ws, r):
        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=r, column=col_idx, value=h)
            cell.font = Font(name="Segoe UI", size=9, bold=True, color=WHITE)
            cell.fill = fill_teal
            cell.alignment = align_wrap_center
            cell.border = border_all
        ws.row_dimensions[r].height = 30

    r = 4
    for phase_num in [1, 2]:
        phase_start_row = r
        write_subheader(ws_mp, r, f"PHASE {phase_num} — {'INITIAL CARRY' if phase_num == 1 else 'MULE DROP'} (DAYS {1 if phase_num == 1 else 4}–{3 if phase_num == 1 else 6})", fill_navy)
        r += 1
        
        phase_day_total_rows = []
        
        for day_entry in menu_data:
            if day_entry["phase"] != phase_num:
                continue
            
            day_num = day_entry["day"]
            date_str = day_entry["date_str"]
            meals = day_entry["meals"]
            
            day_start_row = r
            write_subheader(ws_mp, r, f"{date_str.upper()} — DAY {day_num} DETAILS", fill_teal)
            r += 1
            
            meal_total_rows = []
            
            for meal_name in ["BREAKFAST", "SNACKS", "LUNCH", "SUPPLEMENT", "DINNER", "DESSERT"]:
                items = meals.get(meal_name, [])
                if not items:
                    if meal_name == "BREAKFAST" and day_num == 1:
                        write_subheader(ws_mp, r, f"  {date_str} — BREAKFAST (✗ NOT CARRIED — eat before leaving trailhead)", fill_light_gray)
                        ws_mp.cell(row=r, column=1).font = font_italic
                        ws_mp.cell(row=r, column=1).alignment = align_center
                        r += 1
                    elif meal_name in ["DINNER", "DESSERT"] and day_num == 6:
                        # Exit day no dinner
                        pass
                    continue
                
                # Write meal subheader
                write_subheader(ws_mp, r, f"  {date_str} — {meal_name}", fill_light_gray)
                ws_mp.cell(row=r, column=1).font = font_bold
                r += 1
                
                # Write table headers
                write_table_headers(ws_mp, r)
                r += 1
                
                meal_items_start = r
                for item in items:
                    ws_mp.cell(row=r, column=1, value=item[0])
                    ws_mp.cell(row=r, column=2, value=item[1])
                    ws_mp.cell(row=r, column=3, value=item[2])
                    ws_mp.cell(row=r, column=4, value=item[3])
                    ws_mp.cell(row=r, column=5, value=item[4])
                    # kcal
                    ws_mp.cell(row=r, column=6, value=f"=B{r}*E{r}")
                    ws_mp.cell(row=r, column=7, value=f"=C{r}*E{r}")
                    ws_mp.cell(row=r, column=8, value=f"=D{r}*E{r}")
                    # price
                    ws_mp.cell(row=r, column=9, value=item[5])
                    # cost
                    ws_mp.cell(row=r, column=10, value=f"=B{r}*I{r}")
                    ws_mp.cell(row=r, column=11, value=f"=C{r}*I{r}")
                    ws_mp.cell(row=r, column=12, value=f"=D{r}*I{r}")
                    # oz
                    ws_mp.cell(row=r, column=13, value=item[6])
                    # oz total
                    ws_mp.cell(row=r, column=14, value=f"=B{r}*M{r}")
                    ws_mp.cell(row=r, column=15, value=f"=C{r}*M{r}")
                    ws_mp.cell(row=r, column=16, value=f"=D{r}*M{r}")
                    ws_mp.cell(row=r, column=17, value=item[7])
                    
                    # Formatting
                    is_nut = item[8]
                    is_star = item[9]
                    fill = fill_allergy if is_nut else (fill_adult if is_star else (fill_alt_row if r % 2 == 1 else PatternFill(fill_type=None)))
                    font = font_allergy if is_nut else (font_adult if is_star else font_regular)
                    
                    for col_c in range(1, 18):
                        cell = ws_mp.cell(row=r, column=col_c)
                        cell.font = font
                        cell.border = border_all
                        if fill.fill_type:
                            cell.fill = fill
                        # alignment
                        if col_c == 1 or col_c == 17:
                            cell.alignment = align_left
                        elif col_c in [2, 3, 4, 5, 13]:
                            cell.alignment = align_center
                        else:
                            cell.alignment = align_right
                        
                        # number formats
                        if col_c in [6, 7, 8]:
                            cell.number_format = '#,##0" kcal"'
                        elif col_c in [9, 10, 11, 12]:
                            cell.number_format = '"$"#,##0.00'
                        elif col_c in [14, 15, 16]:
                            cell.number_format = '0.0" oz"'
                    r += 1
                
                meal_items_end = r - 1
                
                # Write meal subtotal
                ws_mp.cell(row=r, column=1, value=f"  {meal_name} Total").font = font_bold
                ws_mp.cell(row=r, column=6, value=f"=SUM(F{meal_items_start}:F{meal_items_end})")
                ws_mp.cell(row=r, column=7, value=f"=SUM(G{meal_items_start}:G{meal_items_end})")
                ws_mp.cell(row=r, column=8, value=f"=SUM(H{meal_items_start}:H{meal_items_end})")
                ws_mp.cell(row=r, column=10, value=f"=SUM(J{meal_items_start}:J{meal_items_end})")
                ws_mp.cell(row=r, column=11, value=f"=SUM(K{meal_items_start}:K{meal_items_end})")
                ws_mp.cell(row=r, column=12, value=f"=SUM(L{meal_items_start}:L{meal_items_end})")
                ws_mp.cell(row=r, column=14, value=f"=SUM(N{meal_items_start}:N{meal_items_end})")
                ws_mp.cell(row=r, column=15, value=f"=SUM(O{meal_items_start}:O{meal_items_end})")
                ws_mp.cell(row=r, column=16, value=f"=SUM(P{meal_items_start}:P{meal_items_end})")
                
                for col_c in range(1, 18):
                    cell = ws_mp.cell(row=r, column=col_c)
                    cell.font = font_bold
                    cell.border = Border(top=thin_side, bottom=thin_side, left=thin_side, right=thin_side)
                    cell.fill = fill_light_gray
                    if col_c in [2, 3, 4, 5, 13]:
                        cell.alignment = align_center
                    elif col_c in [6, 7, 8]:
                        cell.number_format = '#,##0" kcal"'
                        cell.alignment = align_right
                    elif col_c in [10, 11, 12]:
                        cell.number_format = '"$"#,##0.00'
                        cell.alignment = align_right
                    elif col_c in [14, 15, 16]:
                        cell.number_format = '0.0" oz"'
                        cell.alignment = align_right
                
                meal_total_rows.append(r)
                r += 1
                
            # Write Day Total
            ws_mp.cell(row=r, column=1, value=f"DAY {day_num} ({date_str}) TOTAL").font = font_bold
            ws_mp.cell(row=r, column=6, value=f"=SUM(" + ",".join([f"F{x}" for x in meal_total_rows]) + ")")
            ws_mp.cell(row=r, column=7, value=f"=SUM(" + ",".join([f"G{x}" for x in meal_total_rows]) + ")")
            ws_mp.cell(row=r, column=8, value=f"=SUM(" + ",".join([f"H{x}" for x in meal_total_rows]) + ")")
            ws_mp.cell(row=r, column=10, value=f"=SUM(" + ",".join([f"J{x}" for x in meal_total_rows]) + ")")
            ws_mp.cell(row=r, column=11, value=f"=SUM(" + ",".join([f"K{x}" for x in meal_total_rows]) + ")")
            ws_mp.cell(row=r, column=12, value=f"=SUM(" + ",".join([f"L{x}" for x in meal_total_rows]) + ")")
            ws_mp.cell(row=r, column=14, value=f"=SUM(" + ",".join([f"N{x}" for x in meal_total_rows]) + ")")
            ws_mp.cell(row=r, column=15, value=f"=SUM(" + ",".join([f"O{x}" for x in meal_total_rows]) + ")")
            ws_mp.cell(row=r, column=16, value=f"=SUM(" + ",".join([f"P{x}" for x in meal_total_rows]) + ")")
            
            for col_c in range(1, 18):
                cell = ws_mp.cell(row=r, column=col_c)
                cell.font = font_bold
                cell.fill = fill_total
                cell.border = Border(top=thin_side, bottom=thin_side, left=thin_side, right=thin_side)
                if col_c in [6, 7, 8]:
                    cell.number_format = '#,##0" kcal"'
                    cell.alignment = align_right
                elif col_c in [10, 11, 12]:
                    cell.number_format = '"$"#,##0.00'
                    cell.alignment = align_right
                elif col_c in [14, 15, 16]:
                    cell.number_format = '0.0" oz"'
                    cell.alignment = align_right
            
            day_total_rows[day_num] = r
            phase_day_total_rows.append(r)
            r += 2  # spacing
            
        # Write Phase Total
        ws_mp.cell(row=r, column=1, value=f"PHASE {phase_num} SUB-TOTAL (per bear can)").font = font_bold
        ws_mp.cell(row=r, column=6, value=f"=SUM(" + ",".join([f"F{x}" for x in phase_day_total_rows]) + ")")
        ws_mp.cell(row=r, column=7, value=f"=SUM(" + ",".join([f"G{x}" for x in phase_day_total_rows]) + ")")
        ws_mp.cell(row=r, column=8, value=f"=SUM(" + ",".join([f"H{x}" for x in phase_day_total_rows]) + ")")
        ws_mp.cell(row=r, column=10, value=f"=SUM(" + ",".join([f"J{x}" for x in phase_day_total_rows]) + ")")
        ws_mp.cell(row=r, column=11, value=f"=SUM(" + ",".join([f"K{x}" for x in phase_day_total_rows]) + ")")
        ws_mp.cell(row=r, column=12, value=f"=SUM(" + ",".join([f"L{x}" for x in phase_day_total_rows]) + ")")
        ws_mp.cell(row=r, column=14, value=f"=SUM(" + ",".join([f"N{x}" for x in phase_day_total_rows]) + ")")
        ws_mp.cell(row=r, column=15, value=f"=SUM(" + ",".join([f"O{x}" for x in phase_day_total_rows]) + ")")
        ws_mp.cell(row=r, column=16, value=f"=SUM(" + ",".join([f"P{x}" for x in phase_day_total_rows]) + ")")
        
        for col_c in range(1, 18):
            cell = ws_mp.cell(row=r, column=col_c)
            cell.font = Font(name="Segoe UI", size=10, bold=True, color=WHITE)
            cell.fill = fill_navy
            cell.border = Border(top=thin_side, bottom=double_side, left=thin_side, right=thin_side)
            if col_c in [6, 7, 8]:
                cell.number_format = '#,##0" kcal"'
                cell.alignment = align_right
            elif col_c in [10, 11, 12]:
                cell.number_format = '"$"#,##0.00'
                cell.alignment = align_right
            elif col_c in [14, 15, 16]:
                cell.number_format = '0.0" oz"'
                cell.alignment = align_right
                
        phase_total_rows[phase_num] = r
        r += 3  # spacing
        
    # Grand Total
    ws_mp.cell(row=r, column=1, value="GRAND TOTAL (6 DAYS PURCHASE TOTAL)").font = font_bold
    ws_mp.cell(row=r, column=6, value=f"=SUM(F{phase_total_rows[1]},F{phase_total_rows[2]})")
    ws_mp.cell(row=r, column=7, value=f"=SUM(G{phase_total_rows[1]},G{phase_total_rows[2]})")
    ws_mp.cell(row=r, column=8, value=f"=SUM(H{phase_total_rows[1]},H{phase_total_rows[2]})")
    ws_mp.cell(row=r, column=10, value=f"=SUM(J{phase_total_rows[1]},J{phase_total_rows[2]})")
    ws_mp.cell(row=r, column=11, value=f"=SUM(K{phase_total_rows[1]},K{phase_total_rows[2]})")
    ws_mp.cell(row=r, column=12, value=f"=SUM(L{phase_total_rows[1]},L{phase_total_rows[2]})")
    ws_mp.cell(row=r, column=14, value=f"=SUM(N{phase_total_rows[1]},N{phase_total_rows[2]})")
    ws_mp.cell(row=r, column=15, value=f"=SUM(O{phase_total_rows[1]},O{phase_total_rows[2]})")
    ws_mp.cell(row=r, column=16, value=f"=SUM(P{phase_total_rows[1]},P{phase_total_rows[2]})")
    
    for col_c in range(1, 18):
        cell = ws_mp.cell(row=r, column=col_c)
        cell.font = Font(name="Segoe UI", size=11, bold=True)
        cell.fill = fill_total
        cell.border = Border(top=thick_top, bottom=double_side, left=thin_side, right=thin_side)
        if col_c in [6, 7, 8]:
            cell.number_format = '#,##0" kcal"'
            cell.alignment = align_right
        elif col_c in [10, 11, 12]:
            cell.number_format = '"$"#,##0.00'
            cell.alignment = align_right
        elif col_c in [14, 15, 16]:
            cell.number_format = '0.0" oz"'
            cell.alignment = align_right
            
    grand_total_row = r

    # --- 4. SHOPPING LIST SHEET (Generate before Overview to resolve row references) ---
    ws_sl = wb.create_sheet(title="Shopping List")
    ws_sl.views.sheetView[0].showGridLines = True
    
    ws_sl["A1"] = "Master Shopping List — 19 People · 6 Bear Cans  |  July 25–30, 2026"
    ws_sl["A1"].font = font_title
    ws_sl["A2"] = "🟡 Yellow = bulk buy (Costco/Amazon)  |  🟠 Orange = adult supplement  |  Group Total = 3*(Scout 3) + 1*(Scout 4) + 2*(Adult 3)"
    ws_sl["A2"].font = font_subtitle
    
    sl_headers = [
        "☐", "Food Item", "Phase 1 Qty\n(Scout 3 Can)", "Phase 1 Qty\n(Scout 4 Can)", "Phase 1 Qty\n(Adult 3 Can)",
        "Phase 2 Qty\n(Scout 3 Can)", "Phase 2 Qty\n(Scout 4 Can)", "Phase 2 Qty\n(Adult 3 Can)",
        "Group Total\n(6 Days Packs)", "$/Unit", "Scout 3 Cost\n(6 Days)", "Scout 4 Cost\n(6 Days)", "Adult 3 Cost\n(6 Days)",
        "Group Total\n$", "oz/unit", "Group\nozs", "Where to Buy"
    ]
    
    for col_idx, h in enumerate(sl_headers, start=1):
        cell = ws_sl.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_wrap_center
        cell.border = border_all
    ws_sl.row_dimensions[3].height = 30
    
    # Item definitions: name, phase1_sc3, phase1_sc4, phase1_ad3, phase2_sc3, phase2_sc4, phase2_ad3, price, oz, notes, is_nut, is_supp, is_bulk
    sl_data = [
        ["MET-Rx Big100 – Super Cookie Crunch", 6, 8, 6, 9, 12, 9, 2.19, 3.53, "Bkfst; ⚠ NUT", True, False, True],
        ["ProBar Meal – Oatmeal Chocolate Chip", 6, 8, 6, 9, 12, 9, 2.99, 3.00, "Bkfst & Lunch; ⚠ NUT (cashews)", True, False, True],
        ["Peak Refuel – Beef Pasta Marinara", 4, 6, 6, 2, 3, 3, 7.49, 3.20, "Dinner (1.5 svgs/pouch)", False, False, False],
        ["Peak Refuel – Beef Stroganoff", 2, 3, 3, 2, 3, 3, 7.49, 3.50, "Dinner (1.5 svgs/pouch)", False, False, False],
        ["Backpacker's Pantry – Crème Brûlée", 4, 4, 4, 2, 2, 2, 8.99, 2.26, "Dessert (shares pouch)", False, False, False],
        ["Justin's Dark Chocolate Peanut Butter Cups", 3, 4, 3, 3, 4, 3, 1.49, 1.41, "Dessert; ⚠ NUT", True, False, False],
        ["Columbus Hard Salami (sliced packet)", 3, 4, 3, 3, 4, 3, 2.99, 2.00, "Lunch; 1/person", False, False, True],
        ["Wasa Crisp'n Light Crackerbread", 3, 4, 3, 3, 4, 3, 0.50, 1.00, "Lunch side", False, False, False],
        ["Chicken of the Sea Shredded Chicken Pouch", 3, 4, 3, 3, 4, 3, 2.29, 2.50, "Lunch; 1/person", False, False, True],
        ["Mission Street Taco Flour Tortillas", 18, 24, 18, 18, 24, 18, 0.25, 0.50, "lunch + dinner wraps", False, False, True],
        ["Tillamook Cheddar Cheese Packets", 12, 16, 12, 12, 16, 12, 0.75, 0.70, "lunch + dinner side", False, False, True],
        ["Justin's Classic PB Squeeze Pack", 6, 8, 6, 6, 8, 6, 1.49, 1.15, "Lunch side; ⚠ NUT", True, False, False],
        ["Idahoan Mashed Potatoes – Butter & Herb", 3, 4, 3, 3, 4, 3, 0.98, 1.00, "Dinner side", False, False, False],
        ["Kind – Peanut Butter Dark Chocolate Bar", 9, 12, 9, 9, 12, 9, 1.04, 1.40, "Snacks; ⚠ NUT", True, False, True],
        ["Kind – Caramel Almond & Peanut Bar", 9, 12, 9, 9, 12, 9, 1.59, 1.40, "Snacks; ⚠ NUT", True, False, False],
        ["Kind – Dark Choc Nuts & Sea Salt Bar", 9, 12, 9, 9, 12, 9, 1.04, 1.40, "Snacks; ⚠ NUT", True, False, True],
        ["Krave Sea Salt Original Beef Jerky", 9, 12, 9, 9, 12, 9, 2.19, 1.48, "Snacks", False, False, True],
        ["Gatorade Powder Pack", 12, 16, 12, 12, 16, 12, 0.38, 1.23, "Snacks", False, False, True],
        ["Clif Bar – Crunchy PB", 6, 8, 6, 6, 8, 6, 1.29, 2.40, "Lunch snack; ⚠ NUT", True, False, True],
        ["Clif Builder's Bar (adult supplement)", 0, 0, 9, 0, 0, 9, 2.17, 2.40, "Adults ★; ⚠ NUT", True, True, True],
        ["Justin's PB Squeeze (adult supplement)", 0, 0, 9, 0, 0, 9, 1.49, 1.15, "Adults ★; ⚠ NUT", True, True, False],
        ["Swiss Miss Hot Chocolate", 30, 40, 0, 30, 40, 0, 0.38, 0.73, "Scouts only", False, False, True],
        ["Peak Refuel – Fudge Brownie Bites", 2, 2, 2, 2, 2, 2, 4.99, 1.50, "Snacks", False, False, False],
        ["Trader Joe's Dried Mango Slices (1 oz)", 3, 4, 3, 3, 4, 3, 0.55, 1.00, "Snacks", False, False, False],
        ["Banana Chips (1 oz)", 3, 4, 3, 3, 4, 3, 0.45, 1.00, "Snacks", False, False, True],
        
        # Nut allergy swaps
        ["88 Acres Seed + Oat Bars (allergy scout)", 0, 0, 0, 0, 0, 0, 2.40, 1.40, "Allergy scout snacks swap", True, False, False],
        ["Once Again Organic Sunflower Butter Squeeze", 0, 0, 0, 0, 0, 0, 1.50, 1.15, "Allergy scout PB swap", True, False, False],
        ["Honey Stinger Gold Waffle (allergy scout)", 0, 0, 0, 0, 0, 0, 1.09, 1.00, "Allergy scout dessert swap", True, False, False]
    ]
    
    sl_start_row = 4
    for idx, item in enumerate(sl_data):
        r = sl_start_row + idx
        shopping_list_items_rows[item[0]] = r
        
        ws_sl.cell(row=r, column=1, value="☐").alignment = align_center
        ws_sl.cell(row=r, column=2, value=item[0]).alignment = align_left
        
        # Phase 1
        ws_sl.cell(row=r, column=3, value=item[1]).alignment = align_center
        ws_sl.cell(row=r, column=4, value=item[2]).alignment = align_center
        ws_sl.cell(row=r, column=5, value=item[3]).alignment = align_center
        
        # Phase 2
        ws_sl.cell(row=r, column=6, value=item[4]).alignment = align_center
        ws_sl.cell(row=r, column=7, value=item[5]).alignment = align_center
        ws_sl.cell(row=r, column=8, value=item[6]).alignment = align_center
        
        # Group Total formula (6 Days): 3*(Phase 1 + Phase 2 Scout 3) + 1*(Phase 1 + Phase 2 Scout 4) + 2*(Phase 1 + Phase 2 Adult 3)
        # Except for allergy swap items where we set group totals manually!
        is_allergy_item = "allergy scout" in item[9].lower() or "88 acres" in item[0].lower() or "once again" in item[0].lower() or "honey stinger" in item[0].lower()
        if is_allergy_item:
            if "88 acres" in item[0].lower():
                # Swap 2 Kind bars per day on Sat, Sun, Mon, Tue, Wed, Thu = 12 swaps!
                # Wait, Day 3 & Day 6 monday snack has Clif PB -> swapped to Clif Choc Chip (not 88 Acres).
                # Day 1 snacks: PB (1) & Caramel (1) -> 2 swaps.
                # Day 2 snacks: Dark Choc Nuts (1) -> 1 swap.
                # Day 3 snacks: none (uses Jerky + Gatorade + Brownies + Clif PB (swapped to Clif Choc Chip)).
                # Day 4 snacks: Dark Choc Nuts (1) -> 1 swap.
                # Day 5 snacks: PB (1) & Caramel (1) -> 2 swaps.
                # Day 6 snacks: none.
                # Total 88 Acres needed = 2 + 1 + 0 + 1 + 2 + 0 = 6 swaps!
                ws_sl.cell(row=r, column=9, value=6)
            elif "once again" in item[0].lower():
                # Swaps PB squeeze on Sun lunch, Mon lunch, Wed lunch, Thu lunch = 4 swaps!
                ws_sl.cell(row=r, column=9, value=4)
            else: # Honey stinger
                # Swaps Sat dessert (Crème Brûlée, 1), Sun dessert (PB Cups, 1), Mon dessert (Crème Brûlée, 1), Tue dessert (PB Cups, 1), Wed dessert (Crème Brûlée, 1) = 5 swaps!
                ws_sl.cell(row=r, column=9, value=5)
        else:
            ws_sl.cell(row=r, column=9, value=f"=3*(C{r}+F{r})+1*(D{r}+G{r})+2*(E{r}+H{r})")
            
        ws_sl.cell(row=r, column=9).alignment = align_center
        
        # $/Unit
        ws_sl.cell(row=r, column=10, value=item[7]).alignment = align_right
        ws_sl.cell(row=r, column=10).number_format = '"$"#,##0.00'
        
        # Cost columns (K, L, M, N)
        ws_sl.cell(row=r, column=11, value=f"=(C{r}+F{r})*J{r}").alignment = align_right
        ws_sl.cell(row=r, column=11).number_format = '"$"#,##0.00'
        ws_sl.cell(row=r, column=12, value=f"=(D{r}+G{r})*J{r}").alignment = align_right
        ws_sl.cell(row=r, column=12).number_format = '"$"#,##0.00'
        ws_sl.cell(row=r, column=13, value=f"=(E{r}+H{r})*J{r}").alignment = align_right
        ws_sl.cell(row=r, column=13).number_format = '"$"#,##0.00'
        ws_sl.cell(row=r, column=14, value=f"=I{r}*J{r}").alignment = align_right
        ws_sl.cell(row=r, column=14).number_format = '"$"#,##0.00'
        
        # Weight columns (O, P)
        ws_sl.cell(row=r, column=15, value=item[8]).alignment = align_center
        ws_sl.cell(row=r, column=16, value=f"=I{r}*O{r}").alignment = align_right
        ws_sl.cell(row=r, column=16).number_format = '0.0" oz"'
        
        # Notes
        ws_sl.cell(row=r, column=17, value=item[9]).alignment = align_left
        
        # Styles
        is_nut = item[10]
        is_supp = item[11]
        is_bulk = item[12]
        
        fill = fill_allergy if is_nut else (fill_adult if is_supp else (fill_alt_row if r % 2 == 1 else PatternFill(fill_type=None)))
        font = font_allergy if is_nut else (font_adult if is_supp else font_regular)
        
        for col_c in range(1, 18):
            cell = ws_sl.cell(row=r, column=col_c)
            cell.font = font
            cell.border = border_all
            if fill.fill_type:
                cell.fill = fill
                
        # Highlight bulk items with a light gray fill on the Item name or notes if desired
        if is_bulk and not is_nut and not is_supp:
            ws_sl.cell(row=r, column=2).fill = fill_light_gray
            
    sl_end_row = r
    
    # Grand Total Row in Shopping List
    r += 1
    ws_sl.cell(row=r, column=2, value="Grand Total").font = font_bold
    ws_sl.cell(row=r, column=11, value=f"=SUM(K4:K{sl_end_row})")
    ws_sl.cell(row=r, column=12, value=f"=SUM(L4:L{sl_end_row})")
    ws_sl.cell(row=r, column=13, value=f"=SUM(M4:M{sl_end_row})")
    ws_sl.cell(row=r, column=14, value=f"=SUM(N4:N{sl_end_row})")
    ws_sl.cell(row=r, column=16, value=f"=SUM(P4:P{sl_end_row})")
    
    for col_c in range(1, 18):
        cell = ws_sl.cell(row=r, column=col_c)
        cell.font = font_bold
        cell.fill = fill_total
        cell.border = Border(top=thin_side, bottom=double_side, left=thin_side, right=thin_side)
        if col_c in [11, 12, 13, 14]:
            cell.number_format = '"$"#,##0.00'
            cell.alignment = align_right
        elif col_c == 16:
            cell.number_format = '0.0" oz"'
            cell.alignment = align_right
            
    shopping_list_totals_row = r

    # --- 1. OVERVIEW SHEET (Generate now since we have the row references) ---
    ws_ov = wb.create_sheet(title="Overview", index=0)
    ws_ov.views.sheetView[0].showGridLines = True
    
    ws_ov["A1"] = "Emigrant Wilderness — Big Sam Loop Crew Meal Plan (19 People)"
    ws_ov["A1"].font = font_title
    ws_ov["A2"] = "July 25–30, 2026  ·  Kennedy Meadows Loop  ·  ~46.9 mi  ·  13 Scouts · 6 Adults  |  3 Scout cans (3 people) · 1 Scout can (4 people) · 2 Adult cans (3 people)"
    ws_ov["A2"].font = font_subtitle
    
    # Canister Structure Table
    ws_ov["A4"] = "BEAR CAN STRUCTURE"
    ws_ov["A4"].font = font_bold
    
    structure_headers = ["Can assignment", "Can 1", "Can 2", "Can 3", "Can 4", "Can 5", "Can 6", "Type", "People", "Notes"]
    for col_idx, header in enumerate(structure_headers, start=1):
        cell = ws_ov.cell(row=5, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = border_all
        
    structure_data = [
        ["Can assignment", "Scout Can 1", "Scout Can 2", "Scout Can 3", "Scout Can 4", "Adult Can 1", "Adult Can 2", "", "", "4 scout cans + 2 adult cans"],
        ["People per can", "3 scouts", "3 scouts", "3 scouts", "4 scouts", "3 adults", "3 adults", "", "", "Scouts: 3, 3, 3, 4 · Adults: 3, 3"],
        ["Hot drink", "Swiss Miss", "Swiss Miss", "Swiss Miss", "Swiss Miss", "own coffee", "own coffee", "", "", "Scouts: Swiss Miss · Adults: own coffee/tea"],
        ["Carry capacity", "3 Days", "3 Days", "3 Days", "3 Days", "3 Days", "3 Days", "", "", "Balance food is mule dropped (Days 4–6)"]
    ]
    
    for r_idx, row_data in enumerate(structure_data, start=6):
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws_ov.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx == 1 else font_regular
            cell.alignment = align_left if c_idx == 1 or c_idx == 10 else align_center
            cell.border = border_all
            if r_idx % 2 == 1:
                cell.fill = fill_alt_row
                
    # Trip Schedule Table
    ws_ov["A11"] = "TRIP SCHEDULE"
    ws_ov["A11"].font = font_bold
    
    schedule_headers = ["Day", "Date", "Route", "Miles", "Gain", "Loss", "Camp", "Meals", "Water", "Notes"]
    for col_idx, header in enumerate(schedule_headers, start=1):
        cell = ws_ov.cell(row=12, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = border_all
        
    schedule_data = [
        ["Day 1", "Sat Jul 25", "Kennedy Meadows → Lower Relief Valley", "6.88 mi", "+2,262 ft", "−579 ft", "Lower Relief Valley", "Snacks · Lunch · Dinner · Dessert", "Reservoir/Creek (filter)", "morales swimming hole side-trip"],
        ["Day 2", "Sun Jul 26", "Lower Relief Valley → Spring Meadow", "5.97 mi", "+1,129 ft", "−509 ft", "Spring Meadow", "Bkfst · Snacks · Lunch · Dinner · Dessert", "Spring/Creek (filter)", "Wire Lake side-trip"],
        ["Day 3", "Mon Jul 27", "Spring Meadow → eastern Emigrant Lake", "8.89 mi", "+1,210 ft", "−1,035 ft", "Emigrant Lake (east)", "Bkfst · Snacks · Lunch · Dinner · Dessert", "Lake (filter)", "morales swimming spot; Buck Lake on route"],
        ["Day 4", "Tue Jul 28", "Emigrant Lake → Grizzly Peak Base Camp", "6.59 mi", "+1,575 ft", "−742 ft", "Grizzly Peak Base Camp", "Bkfst · Snacks · Lunch · Dinner · Dessert", "Lake (filter)", "Mule drop refilled; cooler alpine camp"],
        ["Day 5", "Wed Jul 29", "Grizzly Peak → Big Sam → Kennedy Lake", "8.46 mi", "+1,428 ft", "−3,238 ft", "South Kennedy Lake", "Bkfst · Snacks · Lunch · Dinner · Dessert", "Creek (filter)", "Climb Big Sam (10,804 ft); long descent"],
        ["Day 6", "Thu Jul 30", "Kennedy Lake → Kennedy Meadows (Exit)", "8.50 mi", "+0 ft", "−1,517 ft", "Exit — Trailhead", "Bkfst · Snacks · Lunch", "Creek/Tap", "Walk out valley; celebrate at trailhead"],
        ["TOTALS", "", "", "51.29 mi", "+7,604 ft", "−8,620 ft", "", "", "", "Totals include optional side-trips"]
    ]
    
    for r_idx, row_data in enumerate(schedule_data, start=13):
        is_total = (row_data[0] == "TOTALS")
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws_ov.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if is_total or c_idx == 1 else font_regular
            cell.alignment = align_left if c_idx == 3 or c_idx == 10 else align_center
            cell.border = border_all
            if is_total:
                cell.fill = fill_total
                cell.border = Border(top=thin_side, bottom=double_side, left=thin_side, right=thin_side)
            elif r_idx % 2 == 1:
                cell.fill = fill_alt_row
                
    # Calorie Overview
    ws_ov["A21"] = "CALORIE OVERVIEW (per person)"
    ws_ov["A21"].font = font_bold
    
    calorie_headers = ["", "Scouts", "Adults", "Notes"]
    for col_idx, header in enumerate(calorie_headers, start=1):
        cell = ws_ov.cell(row=22, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_teal
        cell.alignment = align_center
        cell.border = border_all
        
    calorie_data = [
        ["Target (full active day)", "~2,750 kcal", "~3,750 kcal", "Target calorie values"],
        ["Sat Jul 25 delivered", f"='Meal Plan (Per Bear Can)'!F{day_total_rows[1]}", f"='Meal Plan (Per Bear Can)'!H{day_total_rows[1]}", "Snacks + Lunch + Dinner + Dessert (no breakfast carried)"],
        ["Sun Jul 26 delivered", f"='Meal Plan (Per Bear Can)'!F{day_total_rows[2]}", f"='Meal Plan (Per Bear Can)'!H{day_total_rows[2]}", "Bkfst + Snacks + Lunch + Dinner + Dessert + Adult Supp"],
        ["Mon Jul 27 delivered", f"='Meal Plan (Per Bear Can)'!F{day_total_rows[3]}", f"='Meal Plan (Per Bear Can)'!H{day_total_rows[3]}", "Bkfst + Snacks + Lunch + Dinner + Dessert"],
        ["Tue Jul 28 delivered", f"='Meal Plan (Per Bear Can)'!F{day_total_rows[4]}", f"='Meal Plan (Per Bear Can)'!H{day_total_rows[4]}", "Bkfst + Snacks + Lunch + Dinner + Dessert"],
        ["Wed Jul 29 delivered", f"='Meal Plan (Per Bear Can)'!F{day_total_rows[5]}", f"='Meal Plan (Per Bear Can)'!H{day_total_rows[5]}", "Bkfst + Snacks + Lunch + Dinner + Dessert + Adult Supp"],
        ["Thu Jul 30 delivered", f"='Meal Plan (Per Bear Can)'!F{day_total_rows[6]}", f"='Meal Plan (Per Bear Can)'!H{day_total_rows[6]}", "Bkfst + Snacks + Lunch (exit day; no dinner)"],
        ["TRIP TOTAL (6 DAYS)", f"='Meal Plan (Per Bear Can)'!F{grand_total_row}", f"='Meal Plan (Per Bear Can)'!H{grand_total_row}", "Total food purchased per person across all 6 days"],
        ["MAX CARRY WEIGHT VOL (3 DAYS)", f"=MAX('Meal Plan (Per Bear Can)'!F{phase_total_rows[1]},'Meal Plan (Per Bear Can)'!F{phase_total_rows[2]})", f"=MAX('Meal Plan (Per Bear Can)'!H{phase_total_rows[1]},'Meal Plan (Per Bear Can)'!H{phase_total_rows[2]})", "Maximum food carried at one time in canisters"]
    ]
    
    for r_idx, row_data in enumerate(calorie_data, start=23):
        is_total = "TOTAL" in row_data[0]
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws_ov.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if is_total or c_idx == 1 else font_regular
            cell.alignment = align_left if c_idx == 1 or c_idx == 4 else align_center
            cell.border = border_all
            if is_total:
                cell.fill = fill_total
                cell.border = Border(top=thin_side, bottom=double_side, left=thin_side, right=thin_side)
            elif r_idx % 2 == 1:
                cell.fill = fill_alt_row
            
            if c_idx in [2, 3] and r_idx > 23:
                cell.number_format = '"~"#,##0" kcal"'
                
    # Budget Overview
    ws_ov["A33"] = "BUDGET & WEIGHT OVERVIEW (6 Days Purchased / 3 Days Carried)"
    ws_ov["A33"].font = font_bold
    
    budget_headers = ["", "Adult Can (3 ppl) (×2)", "Scout Can (3 ppl) (×3)", "Scout Can (4 ppl) (×1)", "Group Total", "Notes"]
    for col_idx, header in enumerate(budget_headers, start=1):
        cell = ws_ov.cell(row=34, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_teal
        cell.alignment = align_center
        cell.border = border_all
        
    budget_data = [
        ["Estimated Purchased Cost (6 days)", f"=SUM('Shopping List'!M$4:M${sl_end_row})", f"=SUM('Shopping List'!K$4:K${sl_end_row})", f"=SUM('Shopping List'!L$4:L${sl_end_row})", f"=SUM('Shopping List'!N$4:N${sl_end_row})", "Full 6 days food cost; bulk prices"],
        ["Per person cost", "=B35/3", "=C35/3", "=D35/4", "=E35/19", "Averaged across group"],
        ["Max Carry Weight (3 days)", f"=SUMPRODUCT('Shopping List'!E$4:E${sl_end_row},'Shopping List'!O$4:O${sl_end_row})/16", f"=SUMPRODUCT('Shopping List'!C$4:C${sl_end_row},'Shopping List'!O$4:O${sl_end_row})/16", f"=SUMPRODUCT('Shopping List'!D$4:D${sl_end_row},'Shopping List'!O$4:O${sl_end_row})/16", f"=SUMPRODUCT(3*'Shopping List'!C$4:C${sl_end_row}+1*'Shopping List'!D$4:D${sl_end_row}+2*'Shopping List'!E$4:E${sl_end_row},'Shopping List'!O$4:O${sl_end_row})/16", "Max carried weight in lbs (Phase 1)"],
        ["Mule Drop Weight (3 days)", f"=SUMPRODUCT('Shopping List'!H$4:H${sl_end_row},'Shopping List'!O$4:O${sl_end_row})/16", f"=SUMPRODUCT('Shopping List'!F$4:F${sl_end_row},'Shopping List'!O$4:O${sl_end_row})/16", f"=SUMPRODUCT('Shopping List'!G$4:G${sl_end_row},'Shopping List'!O$4:O${sl_end_row})/16", f"=SUMPRODUCT(3*'Shopping List'!F$4:F${sl_end_row}+1*'Shopping List'!G$4:G${sl_end_row}+2*'Shopping List'!H$4:H${sl_end_row},'Shopping List'!O$4:O${sl_end_row})/16", "Mule drop weight in lbs (Phase 2)"]
    ]
    
    for r_idx, row_data in enumerate(budget_data, start=35):
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws_ov.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx == 1 or c_idx == 5 else font_regular
            cell.alignment = align_left if c_idx == 1 or c_idx == 6 else align_center
            cell.border = border_all
            if r_idx % 2 == 1:
                cell.fill = fill_alt_row
                
            # Formatting
            if r_idx in [35, 36] and c_idx in [2, 3, 4, 5]:
                cell.number_format = '"$"#,##0.00'
            elif r_idx in [37, 38] and c_idx in [2, 3, 4, 5]:
                cell.number_format = '0.0" lbs"'
                
    # Water & Safety notes (re-align for Emigrant Wilderness)
    ws_ov["A40"] = "WATER & SAFETY INFORMATION"
    ws_ov["A40"].font = font_bold
    
    water_headers = ["Location", "Source", "Reliability", "Required Action"]
    for col_idx, header in enumerate(water_headers, start=1):
        cell = ws_ov.cell(row=41, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = border_all
        
    water_data = [
        ["Kennedy Meadows Trailhead", "Pack Station tap", "Reliable", "Fill all bottles. Carry 2-3 L per person."],
        ["Relief Reservoir", "Reservoir outlet", "Reliable", "Morale boost lunch stop. Filter all water."],
        ["Lower Relief Valley Camp", "Creek / Spring", "Reliable", "morales swimming spot side trip nearby."],
        ["Spring Meadow Camp", "Meadow stream", "Reliable", "Morale boost meadow camp. Filter water."],
        ["Emigrant Lake Camp", "Lake water", "Reliable", "Marquee camp. Swim spot. Filter before drinking."],
        ["Grizzly Peak Base Camp", "Alpine lake", "Reliable", "High alpine water. Cooler nights. Filter."],
        ["Kennedy Lake Camp", "Kennedy Creek", "Reliable", "morales swimming spot. Filter."],
        ["Kennedy Meadows Exit", "Pack Station tap", "Reliable", "Morale boost finish. Walk out valley."]
    ]
    
    for r_idx, row_data in enumerate(water_data, start=42):
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws_ov.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_regular
            cell.alignment = align_left
            cell.border = border_all
            if r_idx % 2 == 1:
                cell.fill = fill_alt_row

    # --- 3. PACKING LIST SHEET ---
    ws_pl = wb.create_sheet(title="Packing List")
    ws_pl.views.sheetView[0].showGridLines = True
    
    ws_pl["A1"] = "EMIGRANT WILDERNESS — BEAR CAN PACKING LIST  ·  July 25–30, 2026"
    ws_pl["A1"].font = font_title
    ws_pl["A2"] = "Check off each item as you pack.   ★ = adult supplement.   Qty = items per bear can type."
    ws_pl["A2"].font = font_subtitle
    
    pl_headers = [
        "Packed", "Meal", "Item", "Scout Can\n(3 ppl) Qty", "Scout Can\n(4 ppl) Qty", "Adult Can\n(3 ppl) Qty", "Notes"
    ]
    
    for col_idx, h in enumerate(pl_headers, start=1):
        cell = ws_pl.cell(row=4, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_wrap_center
        cell.border = border_all
    ws_pl.row_dimensions[4].height = 28
    
    pl_rows = [
        # --- PHASE 1 ---
        ("PHASE 1 — INITIAL CARRY (DAYS 1–3) — PACK AT TRAILHEAD", "", "", "", "", "", "", True),
        ("SATURDAY  JUL 25 — DAY 1", "", "", "", "", "", "", True),
        ("", "SNACKS", "Kind – Peanut Butter Dark Chocolate Bar", 3, 4, 3, "hip-belt pocket; ⚠ NUT", False),
        ("", "SNACKS", "Kind – Caramel Almond & Peanut Bar", 3, 4, 3, "hip-belt pocket; ⚠ NUT", False),
        ("", "SNACKS", "Krave Sea Salt Original Beef Jerky", 3, 4, 3, "hip-belt pocket; nut-free", False),
        ("", "SNACKS", "Gatorade Powder Pack", 3, 4, 3, "hip-belt pocket; electrolytes", False),
        ("", "SNACKS", "Trader Joe's Dried Mango Slices (1 oz)", 3, 4, 3, "hip-belt pocket; nut-free", False),
        ("", "LUNCH", "ProBar Meal – Oatmeal Chocolate Chip", 3, 4, 3, "no cook; ⚠ NUT", False),
        ("", "DINNER", "Peak Refuel – Beef Pasta Marinara", 2, 3, 3, "hot · 1.5 svgs/pouch", False),
        ("", "DINNER", "Flour Tortillas", 6, 8, 6, "dinner wraps + side", False),
        ("", "DINNER", "Tillamook Cheddar Cheese Packets", 6, 8, 6, "dinner side", False),
        ("", "DINNER", "Swiss Miss Hot Chocolate", 6, 8, 0, "scouts only", False),
        ("", "DESSERT", "Backpacker's Pantry Crème Brûlée", 2, 2, 2, "hot · buddy-pair shares pouch", False),
        ("", "DESSERT", "Swiss Miss Hot Chocolate", 6, 8, 0, "scouts only", False),
        
        ("SUNDAY  JUL 26 — DAY 2", "", "", "", "", "", "", True),
        ("", "BREAKFAST", "MET-Rx Big100 – Super Cookie Crunch", 3, 4, 3, "grab & go; ⚠ NUT", False),
        ("", "BREAKFAST", "Swiss Miss Hot Chocolate", 6, 8, 0, "scouts only", False),
        ("", "BREAKFAST", "Adults: own coffee / tea", 0, 0, 3, "adults only ★", False),
        ("", "SNACKS", "Krave Sea Salt Original Beef Jerky", 3, 4, 3, "hip-belt pocket; nut-free", False),
        ("", "SNACKS", "Kind – Dark Choc Nuts & Sea Salt Bar", 3, 4, 3, "hip-belt pocket; ⚠ NUT", False),
        ("", "SNACKS", "Gatorade Powder Pack", 6, 8, 6, "×2/person; climb day", False),
        ("", "SNACKS", "Banana Chips (1 oz)", 3, 4, 3, "nut-free", False),
        ("", "LUNCH", "Chicken of the Sea Shredded Chicken Pouch", 3, 4, 3, "no cook; nut-free", False),
        ("", "LUNCH", "Flour Tortillas", 6, 8, 6, "lunch wraps", False),
        ("", "LUNCH", "Tillamook Cheddar Cheese Packets", 3, 4, 3, "lunch side", False),
        ("", "LUNCH", "Clif Bar – Crunchy PB", 3, 4, 3, "lunch snack; ⚠ NUT", False),
        ("", "LUNCH", "Justin's Classic PB Squeeze Pack", 3, 4, 3, "lunch side; ⚠ NUT", False),
        ("", "SUPPLEMENT", "Clif Builder's Bar (adult supplement)", 0, 0, 3, "pre-dinner; adults only ★; ⚠ NUT", False),
        ("", "SUPPLEMENT", "Justin's PB Squeeze (adult supplement)", 0, 0, 3, "any time; adults only ★; ⚠ NUT", False),
        ("", "DINNER", "Peak Refuel – Beef Stroganoff", 2, 3, 3, "hot · 1.5 svgs/pouch", False),
        ("", "DINNER", "Flour Tortillas", 6, 8, 6, "dinner wraps", False),
        ("", "DINNER", "Idahoan Mashed Potatoes", 3, 4, 3, "dinner side", False),
        ("", "DINNER", "Swiss Miss Hot Chocolate", 6, 8, 0, "scouts only", False),
        ("", "DESSERT", "Justin's Dark Choc PB Cups", 3, 4, 3, "1 packet/person; ⚠ NUT", False),
        ("", "DESSERT", "Swiss Miss Hot Chocolate", 6, 8, 0, "scouts only", False),
        
        ("MONDAY  JUL 27 — DAY 3", "", "", "", "", "", "", True),
        ("", "BREAKFAST", "ProBar Meal – Oatmeal Chocolate Chip", 3, 4, 3, "grab & go; ⚠ NUT (cashews)", False),
        ("", "BREAKFAST", "Swiss Miss Hot Chocolate", 6, 8, 0, "scouts only", False),
        ("", "BREAKFAST", "Adults: own coffee / tea", 0, 0, 3, "adults only ★", False),
        ("", "SNACKS", "Krave Sea Salt Original Beef Jerky", 3, 4, 3, "hip-belt pocket; nut-free", False),
        ("", "SNACKS", "Clif Bar – Crunchy PB", 3, 4, 3, "hip-belt pocket; ⚠ NUT", False),
        ("", "SNACKS", "Gatorade Powder Pack", 3, 4, 3, "hip-belt pocket; electrolytes", False),
        ("", "SNACKS", "Peak Refuel – Fudge Brownie Bites", 2, 2, 2, "morales afternoon snack; nut-free", False),
        ("", "LUNCH", "Columbus Hard Salami (packet)", 3, 4, 3, "no cook; nut-free", False),
        ("", "LUNCH", "Wasa Crispbread", 3, 4, 3, "no cook; nut-free", False),
        ("", "LUNCH", "Flour Tortillas", 6, 8, 6, "lunch wraps", False),
        ("", "LUNCH", "Tillamook Cheddar Cheese Packets", 3, 4, 3, "lunch side", False),
        ("", "LUNCH", "Justin's Classic PB Squeeze Pack", 3, 4, 3, "lunch side; ⚠ NUT", False),
        ("", "DINNER", "Peak Refuel – Beef Pasta Marinara", 2, 3, 3, "hot · 1.5 svgs/pouch", False),
        ("", "DINNER", "Flour Tortillas", 6, 8, 6, "dinner wraps + side", False),
        ("", "DINNER", "Tillamook Cheddar Cheese Packets", 6, 8, 6, "dinner side", False),
        ("", "DINNER", "Swiss Miss Hot Chocolate", 6, 8, 0, "scouts only", False),
        ("", "DESSERT", "Backpacker's Pantry Crème Brûlée", 2, 2, 2, "hot · buddy-pair shares pouch", False),
        ("", "DESSERT", "Swiss Miss Hot Chocolate", 6, 8, 0, "scouts only", False),
        
        # --- PHASE 2 ---
        ("PHASE 2 — MULE DROP (DAYS 4–6) — REFILL CANS IN FIELD", "", "", "", "", "", "", True),
        ("TUESDAY  JUL 28 — DAY 4", "", "", "", "", "", "", True),
        ("", "BREAKFAST", "MET-Rx Big100 – Super Cookie Crunch", 3, 4, 3, "grab & go; ⚠ NUT", False),
        ("", "BREAKFAST", "Swiss Miss Hot Chocolate", 6, 8, 0, "scouts only", False),
        ("", "BREAKFAST", "Adults: own coffee / tea", 0, 0, 3, "adults only ★", False),
        ("", "SNACKS", "Krave Sea Salt Original Beef Jerky", 3, 4, 3, "hip-belt pocket; nut-free", False),
        ("", "SNACKS", "Kind – Dark Choc Nuts & Sea Salt Bar", 3, 4, 3, "hip-belt pocket; ⚠ NUT", False),
        ("", "SNACKS", "Gatorade Powder Pack", 6, 8, 6, "×2/person; climb day", False),
        ("", "SNACKS", "Banana Chips (1 oz)", 3, 4, 3, "nut-free", False),
        ("", "LUNCH", "ProBar Meal – Oatmeal Chocolate Chip", 3, 4, 3, "no cook; ⚠ NUT", False),
        ("", "DINNER", "Peak Refuel – Beef Stroganoff", 2, 3, 3, "hot · 1.5 svgs/pouch", False),
        ("", "DINNER", "Flour Tortillas", 6, 8, 6, "dinner wraps", False),
        ("", "DINNER", "Idahoan Mashed Potatoes", 3, 4, 3, "dinner side", False),
        ("", "DINNER", "Swiss Miss Hot Chocolate", 6, 8, 0, "scouts only", False),
        ("", "DESSERT", "Justin's Dark Choc PB Cups", 3, 4, 3, "1 packet/person; ⚠ NUT", False),
        ("", "DESSERT", "Swiss Miss Hot Chocolate", 6, 8, 0, "scouts only", False),
        
        ("WEDNESDAY  JUL 29 — DAY 5", "", "", "", "", "", "", True),
        ("", "BREAKFAST", "ProBar Meal – Oatmeal Chocolate Chip", 3, 4, 3, "grab & go; ⚠ NUT (cashews)", False),
        ("", "BREAKFAST", "Swiss Miss Hot Chocolate", 6, 8, 0, "scouts only", False),
        ("", "BREAKFAST", "Adults: own coffee / tea", 0, 0, 3, "adults only ★", False),
        ("", "SNACKS", "Kind – Peanut Butter Dark Chocolate Bar", 3, 4, 3, "hip-belt pocket; ⚠ NUT", False),
        ("", "SNACKS", "Kind – Caramel Almond & Peanut Bar", 3, 4, 3, "hip-belt pocket; ⚠ NUT", False),
        ("", "SNACKS", "Krave Sea Salt Original Beef Jerky", 3, 4, 3, "hip-belt pocket; nut-free", False),
        ("", "SNACKS", "Gatorade Powder Pack", 3, 4, 3, "hip-belt pocket; electrolytes", False),
        ("", "SNACKS", "Trader Joe's Dried Mango Slices (1 oz)", 3, 4, 3, "hip-belt pocket; nut-free", False),
        ("", "LUNCH", "Chicken of the Sea Shredded Chicken Pouch", 3, 4, 3, "no cook; nut-free", False),
        ("", "LUNCH", "Flour Tortillas", 6, 8, 6, "lunch wraps", False),
        ("", "LUNCH", "Tillamook Cheddar Cheese Packets", 3, 4, 3, "lunch side", False),
        ("", "LUNCH", "Clif Bar – Crunchy PB", 3, 4, 3, "lunch snack; ⚠ NUT", False),
        ("", "LUNCH", "Justin's Classic PB Squeeze Pack", 3, 4, 3, "lunch side; ⚠ NUT", False),
        ("", "SUPPLEMENT", "Clif Builder's Bar (adult supplement)", 0, 0, 3, "pre-dinner; adults only ★; ⚠ NUT", False),
        ("", "SUPPLEMENT", "Justin's PB Squeeze (adult supplement)", 0, 0, 3, "any time; adults only ★; ⚠ NUT", False),
        ("", "DINNER", "Peak Refuel – Beef Pasta Marinara", 2, 3, 3, "hot · 1.5 svgs/pouch", False),
        ("", "DINNER", "Flour Tortillas", 6, 8, 6, "dinner wraps + side", False),
        ("", "DINNER", "Tillamook Cheddar Cheese Packets", 6, 8, 6, "dinner side", False),
        ("", "DINNER", "Swiss Miss Hot Chocolate", 6, 8, 0, "scouts only", False),
        ("", "DESSERT", "Backpacker's Pantry Crème Brûlée", 2, 2, 2, "hot · buddy-pair shares pouch", False),
        ("", "DESSERT", "Swiss Miss Hot Chocolate", 6, 8, 0, "scouts only", False),
        
        ("THURSDAY  JUL 30 — DAY 6", "", "", "", "", "", "", True),
        ("", "BREAKFAST", "MET-Rx Big100 – Super Cookie Crunch", 3, 4, 3, "grab & go; ⚠ NUT", False),
        ("", "BREAKFAST", "Swiss Miss Hot Chocolate", 6, 8, 0, "scouts only", False),
        ("", "BREAKFAST", "Adults: own coffee / tea", 0, 0, 3, "adults only ★", False),
        ("", "SNACKS", "Krave Sea Salt Original Beef Jerky", 3, 4, 3, "hip-belt pocket; nut-free", False),
        ("", "SNACKS", "Clif Bar – Crunchy PB", 3, 4, 3, "hip-belt pocket; ⚠ NUT", False),
        ("", "SNACKS", "Gatorade Powder Pack", 3, 4, 3, "hip-belt pocket; electrolytes", False),
        ("", "SNACKS", "Peak Refuel – Fudge Brownie Bites", 2, 2, 2, "morales afternoon snack; nut-free", False),
        ("", "LUNCH", "Columbus Hard Salami (packet)", 3, 4, 3, "no cook; nut-free", False),
        ("", "LUNCH", "Wasa Crispbread", 3, 4, 3, "no cook; nut-free", False),
        ("", "LUNCH", "Flour Tortillas", 6, 8, 6, "lunch wraps", False),
        ("", "LUNCH", "Tillamook Cheddar Cheese Packets", 3, 4, 3, "lunch side", False),
        ("", "LUNCH", "Justin's Classic PB Squeeze Pack", 3, 4, 3, "lunch side; ⚠ NUT", False)
    ]
    
    pl_row = 5
    for item in pl_rows:
        is_header = item[7]
        if is_header:
            ws_pl.merge_cells(start_row=pl_row, start_column=1, end_row=pl_row, end_column=7)
            cell = ws_pl.cell(row=pl_row, column=1, value=item[0])
            if "PHASE" in item[0]:
                cell.font = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
                cell.fill = fill_navy
                ws_pl.row_dimensions[pl_row].height = 24
            else:
                cell.font = font_subheader
                cell.fill = fill_teal
                ws_pl.row_dimensions[pl_row].height = 22
            cell.alignment = align_left
        else:
            ws_pl.cell(row=pl_row, column=1, value="☐").alignment = align_center
            ws_pl.cell(row=pl_row, column=2, value=item[1]).alignment = align_left
            ws_pl.cell(row=pl_row, column=3, value=item[2]).alignment = align_left
            ws_pl.cell(row=pl_row, column=4, value=item[3]).alignment = align_center
            ws_pl.cell(row=pl_row, column=5, value=item[4]).alignment = align_center
            ws_pl.cell(row=pl_row, column=6, value=item[5]).alignment = align_center
            ws_pl.cell(row=pl_row, column=7, value=item[6]).alignment = align_left
            
            is_nut = "⚠ NUT" in item[6]
            is_star = "★" in item[6]
            
            fill = fill_allergy if is_nut else (fill_adult if is_star else (fill_alt_row if pl_row % 2 == 1 else PatternFill(fill_type=None)))
            font = font_allergy if is_nut else (font_adult if is_star else font_regular)
            
            for col_c in range(1, 8):
                cell = ws_pl.cell(row=pl_row, column=col_c)
                cell.font = font
                cell.border = border_all
                if fill.fill_type:
                    cell.fill = fill
                    
            ws_pl.cell(row=pl_row, column=1).font = Font(name="Segoe UI", size=10, bold=True)
            
        pl_row += 1

    # --- 5. NUT ALLERGY — SCOUT PACK SHEET ---
    ws_na = wb.create_sheet(title="Nut Allergy — Scout Pack")
    ws_na.views.sheetView[0].showGridLines = True
    
    ws_na["A1"] = "Nut Allergy Accommodation — Individual Scout Carry Pack (6 Days)"
    ws_na["A1"].font = font_title
    ws_na["A2"] = "Pre-pack all items in a labelled zip-lock or dry bag for this scout. They carry it personally and swap at each nut-item moment. Main canisters are UNCHANGED."
    ws_na["A2"].font = font_subtitle
    
    na_headers = [
        "Day", "Meal", "Remove (has nuts)", "Nuts in it", "Replace with", "Why safe", "kcal", "Where to buy"
    ]
    
    for col_idx, h in enumerate(na_headers, start=1):
        cell = ws_na.cell(row=4, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = border_all
    ws_na.row_dimensions[4].height = 24
    
    na_data = [
        ["Sat Jul 25", "Snacks", "Kind PB Dark Choc Bar", "Peanuts, almonds", "88 Acres Dark Choc Sea Salt Seed + Oat Bar", "Dedicated nut-free bakery. Free from top 9.", 150, "Amazon / 88acres.com"],
        ["Sat Jul 25", "Snacks", "Kind – Caramel Almond & Peanut Bar", "Almonds, peanuts", "88 Acres Cinnamon Maple Seed + Oat Bar", "Same nut-free facility.", 150, "Amazon / 88acres.com"],
        ["Sat Jul 25", "Dessert", "BP Crème Brûlée", "Check label - may contain cashews", "Honey Stinger Gold Waffle (contingency)", "Honey Stinger: simple ingredients, no nuts.", 160, "REI / Amazon"],
        ["Sun Jul 26", "Breakfast", "MET-Rx Big100 Super Cookie Crunch", "Peanuts, almonds, tree nuts", "No Nuts! Chocolate Chip Protein Bar", "Certified 100% nut-free facility. 12g protein.", 185, "Amazon / gononuts.com"],
        ["Sun Jul 26", "Snacks", "Kind – Dark Choc Nuts & Sea Salt Bar", "Almonds, peanuts, cashews", "88 Acres bar (any flavor)", "Same nut-free facility.", 150, "Amazon / 88acres.com"],
        ["Sun Jul 26", "Lunch", "Clif Bar – Crunchy PB", "Peanuts", "Clif Bar – Chocolate Chip", "Clif Chocolate Chip: no nuts in ingredients.", 250, "Add to main Costco order"],
        ["Sun Jul 26", "Lunch", "Justin's Classic PB Squeeze Pack", "Peanuts", "Once Again Sunflower Butter Squeeze Pack", "Peanut-free facility. Free from top 9.", 200, "Amazon"],
        ["Sun Jul 26", "Dessert", "Justin's Dark Choc PB Cups", "Peanuts, almonds", "Honey Stinger Gold Waffle or Enjoy Life Choc", "No nuts. Enjoy Life: top-14 free.", 160, "REI / Amazon"],
        ["Mon Jul 27", "Breakfast", "ProBar Meal – Oatmeal Chocolate Chip", "Cashews", "No Nuts! Cinnamon Roll or Caramel Mocha", "Same certified nut-free facility.", 185, "Amazon / gononuts.com"],
        ["Mon Jul 27", "Snacks", "Clif Bar – Crunchy PB", "Peanuts", "Clif Bar – Chocolate Chip", "Clif Chocolate Chip: no nuts in ingredients.", 250, "Add to main Costco order"],
        ["Mon Jul 27", "Lunch", "Justin's Classic PB Squeeze Pack", "Peanuts", "Once Again Sunflower Butter Squeeze Pack", "Peanut-free facility. Free from top 9.", 200, "Amazon"],
        ["Tue Jul 28", "Breakfast", "MET-Rx Big100 Super Cookie Crunch", "Peanuts, almonds, tree nuts", "No Nuts! Chocolate Chip Protein Bar", "Certified 100% nut-free facility.", 185, "Amazon"],
        ["Tue Jul 28", "Snacks", "Kind – Dark Choc Nuts & Sea Salt Bar", "Almonds, peanuts, cashews", "88 Acres bar (any flavor)", "Same nut-free facility.", 150, "Amazon"],
        ["Tue Jul 28", "Dessert", "Justin's Dark Choc PB Cups", "Peanuts, almonds", "Honey Stinger Gold Waffle", "No nuts.", 160, "Amazon"],
        ["Wed Jul 29", "Breakfast", "ProBar Meal – Oatmeal Chocolate Chip", "Cashews", "No Nuts! Cinnamon Roll bar", "Same certified nut-free facility.", 185, "Amazon"],
        ["Wed Jul 29", "Snacks", "Kind PB Dark Choc Bar", "Peanuts, almonds", "88 Acres Dark Choc Sea Salt Seed + Oat Bar", "Dedicated nut-free bakery.", 150, "Amazon"],
        ["Wed Jul 29", "Snacks", "Kind – Caramel Almond & Peanut Bar", "Almonds, peanuts", "88 Acres Cinnamon Maple Seed + Oat Bar", "Same nut-free facility.", 150, "Amazon"],
        ["Wed Jul 29", "Lunch", "Clif Bar – Crunchy PB", "Peanuts", "Clif Bar – Chocolate Chip", "Clif Chocolate Chip: no nuts in ingredients.", 250, "Costco"],
        ["Wed Jul 29", "Lunch", "Justin's Classic PB Squeeze Pack", "Peanuts", "Once Again Sunflower Butter Squeeze Pack", "Peanut-free facility.", 200, "Amazon"],
        ["Wed Jul 29", "Dessert", "BP Crème Brûlée", "Check label - may contain cashews", "Honey Stinger Gold Waffle", "No nuts.", 160, "REI / Amazon"],
        ["Thu Jul 30", "Breakfast", "MET-Rx Big100 Super Cookie Crunch", "Peanuts, almonds, tree nuts", "No Nuts! Chocolate Chip Protein Bar", "Certified 100% nut-free facility.", 185, "Amazon"],
        ["Thu Jul 30", "Snacks", "Clif Bar – Crunchy PB", "Peanuts", "Clif Bar – Chocolate Chip", "Clif Chocolate Chip: no nuts in ingredients.", 250, "Costco"],
        ["Thu Jul 30", "Lunch", "Justin's Classic PB Squeeze Pack", "Peanuts", "Once Again Sunflower Butter Squeeze Pack", "Peanut-free facility.", 200, "Amazon"]
    ]
    
    for idx, item in enumerate(na_data):
        r = 5 + idx
        for c in range(1, 9):
            cell = ws_na.cell(row=r, column=c, value=item[c-1])
            cell.font = font_regular
            cell.border = border_all
            if r % 2 == 1:
                cell.fill = fill_alt_row
            
            if c in [1, 2, 7]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
                
            if c == 7:
                cell.number_format = '#,##0" kcal"'

    # --- AUTO-FIT COLUMN WIDTHS & SET ALIGNMENTS ---
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                # If cell is merged, ignore its width calculation to avoid stretching columns
                if type(cell).__name__ == 'MergedCell':
                    continue
                # Split by newline and check max length of any line
                lines = val_str.split('\n')
                for line in lines:
                    if len(line) > max_len:
                        max_len = len(line)
            # overview and other title columns don't stretch column A
            if ws.title == "Overview" and col_letter == "A":
                ws.column_dimensions[col_letter].width = 32
            else:
                ws.column_dimensions[col_letter].width = max(max_len + 3, 10)
                
    wb.save("spring_hike_menu_19_people.xlsx")
    print("Spreadsheet spring_hike_menu_19_people.xlsx generated successfully!")

if __name__ == "__main__":
    create_big_sam_menu()
