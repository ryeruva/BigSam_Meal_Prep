import openpyxl

def run_verification():
    wb = openpyxl.load_workbook("spring_hike_menu_19_people.xlsx", data_only=False)
    
    # 1. Verify Overview sheet canister assignment and counts
    ws_ov = wb["Overview"]
    can1_ppl = ws_ov["B7"].value
    can2_ppl = ws_ov["C7"].value
    can3_ppl = ws_ov["D7"].value
    can4_ppl = ws_ov["E7"].value
    can5_ppl = ws_ov["F7"].value
    can6_ppl = ws_ov["G7"].value
    
    print("Overview Canister Configuration:")
    print(f"Can 1: {can1_ppl}")
    print(f"Can 2: {can2_ppl}")
    print(f"Can 3: {can3_ppl}")
    print(f"Can 4: {can4_ppl}")
    print(f"Can 5: {can5_ppl}")
    print(f"Can 6: {can6_ppl}")
    
    # 2. Check Shopping List sheet formulas and row alignments
    ws_sl = wb["Shopping List"]
    errors = 0
    
    # Check rows 4 to 28
    for r in range(4, 29):
        item_name = ws_sl.cell(row=r, column=2).value
        qty_sc3 = ws_sl.cell(row=r, column=3).value
        qty_sc4 = ws_sl.cell(row=r, column=4).value
        qty_ad3 = ws_sl.cell(row=r, column=5).value
        group_total = ws_sl.cell(row=r, column=6).value
        price = ws_sl.cell(row=r, column=7).value
        cost_sc3 = ws_sl.cell(row=r, column=8).value
        cost_sc4 = ws_sl.cell(row=r, column=9).value
        cost_ad3 = ws_sl.cell(row=r, column=10).value
        cost_group = ws_sl.cell(row=r, column=11).value
        weight_oz = ws_sl.cell(row=r, column=12).value
        weight_group = ws_sl.cell(row=r, column=13).value
        
        # Verify Group Total formula
        is_allergy_item = "allergy scout" in ws_sl.cell(row=r, column=14).value.lower() or "88 acres" in item_name.lower() or "once again" in item_name.lower() or "honey stinger" in item_name.lower()
        if not is_allergy_item:
            expected_group_total_formula = f"=3*C{r}+1*D{r}+2*E{r}"
            if group_total != expected_group_total_formula:
                print(f"Row {r} Error: Group total formula is {group_total}, expected {expected_group_total_formula}")
                errors += 1
                
        # Verify Cost Formulas row alignment
        expected_cost_sc3 = f"=C{r}*G{r}"
        expected_cost_sc4 = f"=D{r}*G{r}"
        expected_cost_ad3 = f"=E{r}*G{r}"
        expected_cost_group = f"=F{r}*G{r}"
        
        if cost_sc3 != expected_cost_sc3:
            print(f"Row {r} Error: Scout 3 Cost formula is {cost_sc3}, expected {expected_cost_sc3}")
            errors += 1
        if cost_sc4 != expected_cost_sc4:
            print(f"Row {r} Error: Scout 4 Cost formula is {cost_sc4}, expected {expected_cost_sc4}")
            errors += 1
        if cost_ad3 != expected_cost_ad3:
            print(f"Row {r} Error: Adult 3 Cost formula is {cost_ad3}, expected {expected_cost_ad3}")
            errors += 1
        if cost_group != expected_cost_group:
            print(f"Row {r} Error: Group Cost formula is {cost_group}, expected {expected_cost_group}")
            errors += 1
            
        # Verify Weight Formula row alignment
        expected_weight_group = f"=F{r}*L{r}"
        if weight_group != expected_weight_group:
            print(f"Row {r} Error: Group Weight formula is {weight_group}, expected {expected_weight_group}")
            errors += 1

    # Check Budget Overview cells in Overview
    cost_adult_formula = ws_ov["B28"].value
    cost_sc3_formula = ws_ov["C28"].value
    cost_sc4_formula = ws_ov["D28"].value
    cost_total_formula = ws_ov["E28"].value
    
    print("\nOverview Budget Formulas:")
    print(f"Adult Can Cost Cell B28: {cost_adult_formula}")
    print(f"Scout 3-ppl Cost Cell C28: {cost_sc3_formula}")
    print(f"Scout 4-ppl Cost Cell D28: {cost_sc4_formula}")
    print(f"Group Total Cost Cell E28: {cost_total_formula}")
    
    # 3. Check Meal Plan Formulas in Meal Plan (Per Bear Can) sheet
    ws_mp = wb["Meal Plan (Per Bear Can)"]
    # Check a few rows to verify kcal and cost formula structure
    print("\nMeal Plan Formula Previews:")
    for check_row in [10, 28, 61, 78]:
        item_name = ws_mp.cell(row=check_row, column=1).value
        sq3_qty = ws_mp.cell(row=check_row, column=2).value
        sq4_qty = ws_mp.cell(row=check_row, column=3).value
        ad3_qty = ws_mp.cell(row=check_row, column=4).value
        sq3_kcal = ws_mp.cell(row=check_row, column=6).value
        sq4_kcal = ws_mp.cell(row=check_row, column=7).value
        ad3_kcal = ws_mp.cell(row=check_row, column=8).value
        print(f"Row {check_row} [{item_name}]: Qty({sq3_qty}, {sq4_qty}, {ad3_qty}) | kcal formulas({sq3_kcal}, {sq4_kcal}, {ad3_kcal})")
        
        # Check kcal calculations
        if sq3_kcal != f"=B{check_row}*E{check_row}":
            print(f"Row {check_row} Error: Scout 3 kcal is {sq3_kcal}, expected =B{check_row}*E{check_row}")
            errors += 1
        if sq4_kcal != f"=C{check_row}*E{check_row}":
            print(f"Row {check_row} Error: Scout 4 kcal is {sq4_kcal}, expected =C{check_row}*E{check_row}")
            errors += 1
        if ad3_kcal != f"=D{check_row}*E{check_row}":
            print(f"Row {check_row} Error: Adult 3 kcal is {ad3_kcal}, expected =D{check_row}*E{check_row}")
            errors += 1
            
    print(f"\nVerification finished. Total errors found: {errors}")
    if errors == 0:
        print("PASSED! Spreadsheet formulas are perfectly aligned and correct.")
    else:
        print("FAILED! Please correct the errors.")

if __name__ == "__main__":
    run_verification()
