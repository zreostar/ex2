import openpyxl
from openpyxl.styles import Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# 定義時間區間
TIME_SLOTS = [
    "08:00~08:50",
    "09:00~09:50",
    "10:00~10:50",
    "11:00~11:50",
    "01:00~01:50",
    "02:00~02:50",
    "03:00~03:50",
    "04:00~04:50",
]
DAYS = ["週一", "週二", "週三", "週四", "週五"]

# 課程資料
courses = [
    {"name": "電子實習", "teacher": "邱聰輝教授", "day": "週一", "start": "08:00", "end": "10:50"},
    {"name": "雲端化無限存取實務", "teacher": "劉恩成教授", "day": "週二", "start": "13:00", "end": "15:50"},
    {"name": "工智慧概論", "teacher": "賴文政教授", "day": "週三", "start": "08:00", "end": "08:50"},
    {"name": "工程數學", "teacher": "郭慶祥教授", "day": "週三", "start": "09:00", "end": "11:00"},
    {"name": "專題製作", "teacher": "祁存廣教授", "day": "週三", "start": "15:00", "end": "16:50"},
    {"name": "英文聽講", "teacher": "吳柏德教授", "day": "週四", "start": "08:00", "end": "09:50"},
    {"name": "自動控制與實習", "teacher": "謝飛虎教授", "day": "週四", "start": "13:00", "end": "16:50"},
    {"name": "工程數學", "teacher": "郭慶祥教授", "day": "週五", "start": "08:00", "end": "08:50"},
    {"name": "人工智慧概論", "teacher": "賴文政教授", "day": "週五", "start": "09:00", "end": "10:50"},
    {"name": "影像處理概論", "teacher": "王柏仁教授", "day": "週五", "start": "13:00", "end": "15:50"},
]

# 時間區間轉換
import datetime

def time_to_slot_idx(time_str):
    t = datetime.datetime.strptime(time_str, "%H:%M")
    for idx, slot in enumerate(TIME_SLOTS):
        start = datetime.datetime.strptime(slot.split("~")[0], "%H:%M")
        end = datetime.datetime.strptime(slot.split("~")[1], "%H:%M")
        if start <= t <= end:
            return idx
    # 若精確對齊區間起點
    for idx, slot in enumerate(TIME_SLOTS):
        if slot.startswith(time_str):
            return idx
    return None

def get_slot_range(start, end):
    start_idx = time_to_slot_idx(start)
    end_idx = time_to_slot_idx(end)
    # end_idx 需包含最後一格
    if end_idx is None:
        # 若 end 剛好是區間結束
        for idx, slot in enumerate(TIME_SLOTS):
            if slot.endswith(end):
                end_idx = idx
                break
    if start_idx is not None and end_idx is not None:
        return list(range(start_idx, end_idx+1))
    return []

# 建立課表矩陣
cell_matrix = [[{"course": None, "teacher": None, "conflict": False} for _ in DAYS] for _ in TIME_SLOTS]
merge_info = {}
conflicts = []

for course in courses:
    day_idx = DAYS.index(course["day"])
    slot_range = get_slot_range(course["start"], course["end"])
    # 檢查衝突
    for slot in slot_range:
        if cell_matrix[slot][day_idx]["course"]:
            cell_matrix[slot][day_idx]["conflict"] = True
            conflicts.append(f"衝突：{course['day']} {TIME_SLOTS[slot]} {course['name']} 與 {cell_matrix[slot][day_idx]['course']}")
        else:
            cell_matrix[slot][day_idx]["course"] = course["name"]
            cell_matrix[slot][day_idx]["teacher"] = course["teacher"]
    # 合併儲存格資訊
    if slot_range:
        merge_info[(slot_range[0], day_idx)] = len(slot_range)

# 建立 Excel
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "課表"

# 欄位標題
ws.cell(row=1, column=1, value="時間")
for i, day in enumerate(DAYS):
    ws.cell(row=1, column=i+2, value=day)

# 時間欄
for i, slot in enumerate(TIME_SLOTS):
    ws.cell(row=i+2, column=1, value=slot)

# 格式設定
center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)
red_fill = PatternFill("solid", fgColor="FF0000")

# 填入課程
for col, day in enumerate(DAYS):
    for row, slot in enumerate(TIME_SLOTS):
        cell = ws.cell(row=row+2, column=col+2)
        info = cell_matrix[row][col]
        if info["course"]:
            cell.value = f"{info['course']}\n{info['teacher']}"
            cell.alignment = center_alignment
            cell.border = border
            # 衝突格底色
            if info["conflict"]:
                cell.fill = red_fill
        else:
            cell.value = ""
            cell.alignment = center_alignment
            cell.border = border

# 合併儲存格
for (row, col), span in merge_info.items():
    if span > 1:
        ws.merge_cells(start_row=row+2, start_column=col+2, end_row=row+2+span-1, end_column=col+2)

# 欄寬
ws.column_dimensions[get_column_letter(1)].width = 13
for i in range(2, 7):
    ws.column_dimensions[get_column_letter(i)].width = 22

# 輸出
wb.save("timetable.xlsx")

# 衝突提示
if conflicts:
    for msg in conflicts:
        print(msg)
else:
    print("課表已生成，無衝突。")
