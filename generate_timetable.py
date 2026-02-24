import openpyxl
from openpyxl.styles import Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# 時間區間設定
TIME_SLOTS = [
    ("08:00", "08:50"),
    ("09:00", "09:50"),
    ("10:00", "10:50"),
    ("11:00", "11:50"),
    ("13:00", "13:50"),
    ("14:00", "14:50"),
    ("15:00", "15:50"),
    ("16:00", "16:50"),
]
TIME_SLOT_STRS = [f"{start}~{end}" for start, end in TIME_SLOTS]
WEEKDAYS = ["週一", "週二", "週三", "週四", "週五"]

# 課程資料
COURSES = [
    {"name": "電子實習", "teacher": "邱聰輝教授", "weekday": "週一", "start": "08:00", "end": "10:50"},
    {"name": "雲端化無限存取實務", "teacher": "劉恩成教授", "weekday": "週二", "start": "13:00", "end": "15:50"},
    {"name": "人工智慧概論", "teacher": "賴文政教授", "weekday": "週三", "start": "08:00", "end": "08:50"},
    {"name": "工程數學", "teacher": "郭慶祥教授", "weekday": "週三", "start": "09:00", "end": "11:00"},
    {"name": "專題製作", "teacher": "祁存廣教授", "weekday": "週三", "start": "15:00", "end": "16:50"},
    {"name": "英文聽講", "teacher": "吳柏德教授", "weekday": "週四", "start": "08:00", "end": "09:50"},
    {"name": "自動控制與實習", "teacher": "謝飛虎教授", "weekday": "週四", "start": "13:00", "end": "16:50"},
    {"name": "工程數學", "teacher": "郭慶祥教授", "weekday": "週五", "start": "08:00", "end": "08:50"},
    {"name": "人工智慧概論", "teacher": "賴文政教授", "weekday": "週五", "start": "09:00", "end": "10:50"},
    {"name": "影像處理概論", "teacher": "王柏仁教授", "weekday": "週五", "start": "13:00", "end": "15:50"},
]

def time_to_index(time_str):
    for idx, (start, end) in enumerate(TIME_SLOTS):
        if time_str == start:
            return idx
    raise ValueError(f"Invalid time: {time_str}")

def get_slot_range(start, end):
    start_idx = time_to_index(start)
    # end: 找到最後一個slot包含end
    for idx, (slot_start, slot_end) in enumerate(TIME_SLOTS):
        if end <= slot_end:
            return start_idx, idx
    return start_idx, len(TIME_SLOTS) - 1

def main():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Timetable"
    # 標題列
    ws.append(["時間"] + WEEKDAYS)
    for i, slot in enumerate(TIME_SLOT_STRS, start=2):
        ws.cell(row=i, column=1, value=slot)
    # 初始化格子內容
    cell_map = {}  # (row, col): (course, rowspan)
    conflict_cells = set()
    for course in COURSES:
        col = WEEKDAYS.index(course["weekday"]) + 2
        start_idx, end_idx = get_slot_range(course["start"], course["end"])
        rowspan = end_idx - start_idx + 1
        # 檢查衝突
        for r in range(start_idx + 2, end_idx + 3):
            if (r, col) in cell_map:
                print(f"衝突：{course['weekday']} {TIME_SLOT_STRS[r-2]} {course['name']} 與 {cell_map[(r, col)][0]['name']}")
                conflict_cells.add((r, col))
        # 填入課程（只在起始格）
        cell_map[(start_idx + 2, col)] = (course, rowspan)
        # 標記合併範圍
        for r in range(start_idx + 2, end_idx + 3):
            if (r, col) not in cell_map:
                cell_map[(r, col)] = (None, 0)
    # 寫入課表
    for (r, c), (course, rowspan) in cell_map.items():
        if course:
            ws.cell(row=r, column=c, value=f"{course['name']}\n/ {course['teacher']}")
            if rowspan > 1:
                ws.merge_cells(start_row=r, start_column=c, end_row=r+rowspan-1, end_column=c)
    # 格式化
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    for r in range(1, len(TIME_SLOTS) + 2):
        for c in range(1, len(WEEKDAYS) + 2):
            cell = ws.cell(row=r, column=c)
            cell.alignment = align
            cell.border = border
            if (r, c) in conflict_cells:
                cell.fill = PatternFill("solid", fgColor="FF0000")
    # 欄寬
    ws.column_dimensions[get_column_letter(1)].width = 13
    for i in range(2, 7):
        ws.column_dimensions[get_column_letter(i)].width = 22
    wb.save("timetable.xlsx")
    print("課表已輸出 timetable.xlsx")

if __name__ == "__main__":
    main()
