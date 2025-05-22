import os
import xlwings as xw

def get_unit_cost(sheet):
    """시트에서 'Unit Cost' 오른쪽 셀 값을 찾고, 없으면 H4 반환."""
    try:
        used_range = sheet.used_range
        values = used_range.value
        if values is None:
            return None
        for row in values:
            if isinstance(row, list):
                for c, cell in enumerate(row):
                    if cell == "Unit Cost":
                        if c + 1 < len(row):
                            return row[c + 1]
        return sheet.range('H4').value
    except Exception:
        return None

def update_summary(book, summary_name="Summary"):
    # 1. Summary 시트가 없으면 추가, 있으면 가져오기
    if summary_name in [s.name for s in book.sheets]:
        summary = book.sheets[summary_name]
    else:
        summary = book.sheets.add(summary_name, before=book.sheets[0])
        summary.range("A1").value = ["Sheet Name", "Unit Cost"]

    # 2. 기존 Summary 데이터 읽기 (시트명 → Unit Cost)
    existing = summary.range("A2").expand("table").value
    if existing is None:
        existing = []
    elif isinstance(existing[0], str):
        existing = [[existing[0], summary.range("B2").value]]
    prev_map = {row[0]: row[1] for row in existing if row and row[0]}

    # 3. 현재 시트 정보 수집
    sheet_infos = []
    for sheet in book.sheets:
        if sheet.name == summary_name:
            continue
        unit_cost = get_unit_cost(sheet)
        sheet_infos.append((sheet.name, unit_cost))

    # 4. 변화 내역 추적 및 정렬
    updated, added, removed = [], [], []
    curr_names = set([s[0] for s in sheet_infos])
    prev_names = set(prev_map.keys())

    for sheet_name, unit_cost in sheet_infos:
        if sheet_name in prev_map:
            if prev_map[sheet_name] != unit_cost:
                updated.append((sheet_name, prev_map[sheet_name], unit_cost))
        else:
            added.append((sheet_name, unit_cost))
    for sheet_name in prev_names - curr_names:
        removed.append((sheet_name, prev_map[sheet_name]))

    # 5. Summary 시트 내용 전체 재작성 (빈 줄 없이)
    summary.range("A2:B1048576").clear_contents()
    if sheet_infos:
        summary.range("A2").value = sheet_infos

    summary.autofit()

    # 6. 콘솔에 변화 내역 출력
    for sheet_name, old, new in updated:
        print(f"  Updated: {sheet_name} | {old} → {new}")
    for sheet_name, val in added:
        print(f"  Added: {sheet_name} | {val}")
    for sheet_name, val in removed:
        print(f"  Removed: {sheet_name} | {val}")

    print("  [Summary Sheet]")
    print("  Sheet Name\tUnit Cost")
    for row in sheet_infos:
        print(f"  {row[0]}\t{row[1]}")

def process_all_excels_in_folder(folder):
    excel_files = [f for f in os.listdir(folder) if f.lower().endswith(('.xlsx', '.xlsm'))]
    print(f"Found {len(excel_files)} Excel files in '{folder}'")
    for filename in excel_files:
        filepath = os.path.join(folder, filename)
        print(f"\nProcessing: {filename}")
        app = xw.App(visible=False)
        try:
            book = app.books.open(filepath)
            update_summary(book)
            book.save()
            book.close()
            print("  [OK] Summary updated.\n")
        except Exception as e:
            print(f"  [FAIL] {filename}: {e}")
        finally:
            app.quit()

if __name__ == "__main__":
    folder = os.path.dirname(os.path.abspath(__file__))  # 현재 스크립트 폴더
    # folder = r"C:\원하는\폴더\경로"  # 또는 경로 직접 지정
    process_all_excels_in_folder(folder)
