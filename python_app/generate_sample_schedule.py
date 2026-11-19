"""
Script phụ trợ: Tạo file Schedule.xlsx mẫu (17 tuần hoặc 21 tuần)
đúng theo cấu trúc bảng tính thực tế để kiểm thử chương trình schedule_splitter.py.
"""

from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


def create_sample_schedule(output_path: str = "Schedule_17_Weeks.xlsx", total_weeks: int = 17, is_afternoon: bool = True) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Schedule"

    start_date = datetime(2026, 8, 31)  # Thứ Hai 31/08/2026
    end_date = start_date + timedelta(days=total_weeks * 7 - 1)

    header_text = (
        f"THỜI KHÓA BIỂU - HỌC KỲ FALL 2026 "
        f"({start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')})\n"
        f"Lớp: PC2204     Phòng học: 107"
    )
    ws["A1"] = header_text
    ws["A1"].font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    ws["A1"].alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 40

    weekdays = ["Thứ hai", "Thứ ba", "Thứ tư", "Thứ năm", "Thứ sáu", "Thứ bảy", "CN"]
    slots = [6, 7, 8, 9, 10] if is_afternoon else [1, 2, 3, 4, 5]

    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    block_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    # Mẫu phân bố môn học theo thứ (0=T2..6=CN) và tiết
    weekly_template = {
        (0, 0): "MAT001", (0, 1): "MAT001",
        (1, 0): "PRO192", (1, 1): "PRO192", (1, 2): "ENG101",
        (2, 0): "CSI104", (2, 1): "CSI104",
        (3, 0): "PRO192", (3, 1): "PRO192", (3, 3): "WEB201",
        (4, 0): "MAT001", (4, 2): "ENG101", (4, 3): "ENG101",
        (5, 0): "WEB201", (5, 1): "WEB201",
    }

    current_row = 3
    current_date = start_date
    week_cursor = 1

    while week_cursor <= total_weeks:
        weeks_in_block = min(3, total_weeks - week_cursor + 1)
        end_week_in_block = week_cursor + weeks_in_block - 1
        label = f"Tuần {week_cursor}-{end_week_in_block}" if weeks_in_block > 1 else f"Tuần {week_cursor}"

        # Dòng 1 của khối: Nhãn khối tuần
        ws.cell(row=current_row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=current_row, column=1).fill = block_fill

        # Dòng 2: Nhãn tuần con
        for w_offset in range(weeks_in_block):
            col_start = 2 + w_offset * 7
            ws.cell(row=current_row + 1, column=col_start, value=f"Tuần {week_cursor + w_offset}").font = Font(bold=True)

        # Dòng 3 & 4: Nhãn Thứ và Ngày trong tháng
        total_days = weeks_in_block * 7
        for d_offset in range(total_days):
            col = 2 + d_offset
            day_date = current_date + timedelta(days=d_offset)
            day_of_week = d_offset % 7

            c_wd = ws.cell(row=current_row + 2, column=col, value=weekdays[day_of_week])
            c_wd.alignment = Alignment(horizontal="center")
            c_wd.font = Font(bold=True, size=9)
            c_wd.border = border

            c_dt = ws.cell(row=current_row + 3, column=col, value=day_date.day)
            c_dt.alignment = Alignment(horizontal="center")
            c_dt.border = border

        # 5 dòng tiết học
        for s_idx, slot_num in enumerate(slots):
            r_slot = current_row + 4 + s_idx
            c_label = ws.cell(row=r_slot, column=1, value=f"Tiết {slot_num}")
            c_label.font = Font(bold=True)
            c_label.border = border

            for d_offset in range(total_days):
                col = 2 + d_offset
                day_date = current_date + timedelta(days=d_offset)
                day_of_week = d_offset % 7
                cell = ws.cell(row=r_slot, column=col)
                cell.alignment = Alignment(horizontal="center")
                cell.border = border

                # Chèn ghi chú nghỉ lễ 2/9 và khai giảng để kiểm thử bộ lọc
                if day_date.month == 9 and day_date.day == 2 and s_idx in (0, 1):
                    cell.value = "Nghỉ lễ 2/9"
                elif week_cursor == 1 and day_of_week == 0 and s_idx == 2:
                    cell.value = "Khai giảng K22"
                elif (day_of_week, s_idx) in weekly_template:
                    cell.value = weekly_template[(day_of_week, s_idx)]

        current_date += timedelta(days=total_days)
        week_cursor += weeks_in_block
        current_row += 10

    ws.column_dimensions["A"].width = 14
    wb.save(output_path)
    print(f"Đã tạo file mẫu thành công: {output_path} ({total_weeks} tuần)")


if __name__ == "__main__":
    create_sample_schedule("Schedule_17_Weeks.xlsx", total_weeks=17, is_afternoon=True)
    create_sample_schedule("Schedule_21_Weeks.xlsx", total_weeks=21, is_afternoon=False)
