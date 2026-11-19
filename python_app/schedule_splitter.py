"""
Chương trình: Tự động bóc tách Thời khóa biểu (Schedule Splitter)
Mô tả:
    - Đọc file thời khóa biểu tổng quát dạng Excel (Schedule.xlsx).
    - Tự động bóc tách thông tin tiêu đề bằng Regex (Học kỳ, Ngày bắt đầu/kết thúc, Lớp, Phòng).
    - Quét động toàn bộ các khối tuần (hỗ trợ mọi số lượng tuần: 15, 17, 21 tuần hoặc bất kỳ).
    - Tự động nhận diện ca học (Ca Sáng / Ca Chiều) và lọc bỏ ghi chú sự kiện / ngày nghỉ.
    - Cho phép người dùng gán tên Giảng viên cho từng môn học trên giao diện Tkinter.
    - Đánh số thứ tự buổi học (Session No.) theo trình tự thời gian và xuất từng file
      <Class>_<Subject>.xlsx được định dạng chuẩn đẹp bằng openpyxl.
"""

from __future__ import annotations

import os
import re
import tkinter as tk
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ==============================================================================
# CÁC LỚP DỮ LIỆU (DATA MODELS)
# ==============================================================================

@dataclass
class ScheduleMetadata:
    """Lưu trữ thông tin chung bóc tách từ phần tiêu đề file Schedule.xlsx."""
    term: str
    date_begin: datetime
    date_end: datetime
    class_name: str
    room: str
    shift_label: str = "Chưa xác định"
    total_weeks: int = 0
    total_blocks: int = 0


@dataclass
class SessionRecord:
    """Đại diện cho một buổi học hợp lệ được trích xuất từ thời khóa biểu."""
    term: str
    date_obj: datetime
    date_str: str
    slot: int
    room: str
    subject: str
    class_name: str
    lecturer: str = ""
    session_no: int = 0


@dataclass
class ParseResult:
    """Kết quả sau khi quét động toàn bộ workbook thời khóa biểu."""
    metadata: ScheduleMetadata
    records: List[SessionRecord] = field(default_factory=list)
    subjects: List[str] = field(default_factory=list)
    ignored_notes: List[Tuple[str, str, int]] = field(default_factory=list)  # (ghi_chú, ngày, tiết)


# ==============================================================================
# XỬ LÝ LOGIC BÓC TÁCH VÀ QUÉT ĐỘNG (CORE PARSER)
# ==============================================================================

VIETNAMESE_DIACRITICS_RE = re.compile(
    r"[àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩ"
    r"òóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ"
    r"ÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴÈÉẸẺẼÊỀẾỆỂỄÌÍỊỈĨ"
    r"ÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠÙÚỤỦŨƯỪỨỰỬỮỲÝỴỶỸĐ]"
)

# Các từ khóa tiếng Anh / ghi chú phổ biến không phải mã môn học
NON_SUBJECT_KEYWORDS = {
    "NGHI", "LE", "HOLIDAY", "OFF", "BREAK", "EVENT",
    "ORIENTATION", "MEETING", "EXAM", "TEST", "FINAL",
    "MIDTERM", "CEREMONY", "OPENING", "CLOSING", "NOTE",
}

WEEK_BLOCK_HEADER_RE = re.compile(r"^Tuần\s+(\d+)(?:\s*[-–]\s*(\d+))?$", re.IGNORECASE)
SUB_WEEK_RE = re.compile(r"^Tuần\s+\d+$", re.IGNORECASE)
SLOT_LABEL_RE = re.compile(r"^Tiết\s*(\d+)$", re.IGNORECASE)
WEEKDAY_MARKER_RE = re.compile(
    r"(thứ\s*(?:[2-7]|hai|ba|tư|năm|sáu|bảy)|chủ\s*nhật|^cn$|^t[2-7]$|mon|tue|wed|thu|fri|sat|sun)",
    re.IGNORECASE,
)


def parse_header_metadata(ws: openpyxl.worksheet.worksheet.Worksheet) -> ScheduleMetadata:
    """
    Đọc phần tiêu đề ở đầu sheet (ưu tiên ô A1 và các dòng đầu trước khối Tuần đầu tiên)
    và dùng Regex để trích xuất Term, Date_Begin, Date_End, Class, Room.
    """
    header_lines: List[str] = []
    for r in range(1, min(6, ws.max_row + 1)):
        row_texts: List[str] = []
        for c in range(1, min(10, ws.max_column + 1)):
            val = ws.cell(row=r, column=c).value
            if val is not None:
                text = str(val).strip()
                if text:
                    # Nếu đã gặp dòng bắt đầu khối tuần thì dừng thu thập tiêu đề
                    if c == 1 and WEEK_BLOCK_HEADER_RE.match(text):
                        break
                    row_texts.append(text)
        if row_texts and WEEK_BLOCK_HEADER_RE.match(str(ws.cell(row=r, column=1).value or "").strip()):
            break
        if row_texts:
            header_lines.append(" ".join(row_texts))

    full_header = "\n".join(header_lines)
    if not full_header.strip():
        raise ValueError("Không tìm thấy nội dung tiêu đề ở dòng đầu tiên (ô A1).")

    # 1. Trích xuất Học kỳ (Term)
    term_match = re.search(
        r"HỌC\s*KỲ\s*:?\s*([A-Za-z0-9\s_-]+?)(?=\s*\()",
        full_header,
        re.IGNORECASE,
    )
    if not term_match:
        raise ValueError("Không thể bóc tách tên Học kỳ (ví dụ: HỌC KỲ FALL 2026) từ tiêu đề.")
    term = re.sub(r"\s+", " ", term_match.group(1)).strip().upper()

    # 2. Trích xuất Ngày bắt đầu & Ngày kết thúc
    date_match = re.search(
        r"\(?\s*(\d{1,2}/\d{1,2}/\d{4})\s*[-–]\s*(\d{1,2}/\d{1,2}/\d{4})\s*\)?",
        full_header,
    )
    if not date_match:
        raise ValueError("Không thể bóc tách khoảng ngày (DD/MM/YYYY - DD/MM/YYYY) từ tiêu đề.")

    try:
        date_begin = datetime.strptime(date_match.group(1), "%d/%m/%Y")
        date_end = datetime.strptime(date_match.group(2), "%d/%m/%Y")
    except ValueError as exc:
        raise ValueError(f"Định dạng ngày không hợp lệ trong tiêu đề: {exc}") from exc

    if date_end < date_begin:
        raise ValueError("Ngày kết thúc học kỳ nhỏ hơn ngày bắt đầu.")

    # 3. Trích xuất Mã lớp (Class)
    class_match = re.search(r"Lớp\s*:\s*([A-Za-z0-9._-]+)", full_header, re.IGNORECASE)
    if not class_match:
        raise ValueError("Không thể bóc tách Mã lớp (ví dụ: Lớp: PC2204) từ tiêu đề.")
    class_name = class_match.group(1).strip().upper()

    # 4. Trích xuất Phòng học (Room)
    room_match = re.search(r"Phòng\s*(?:học)?\s*:\s*([A-Za-z0-9._-]+)", full_header, re.IGNORECASE)
    if not room_match:
        raise ValueError("Không thể bóc tách Phòng học (ví dụ: Phòng học: 107) từ tiêu đề.")
    room = room_match.group(1).strip()

    return ScheduleMetadata(
        term=term,
        date_begin=date_begin,
        date_end=date_end,
        class_name=class_name,
        room=room,
    )


def normalize_and_validate_subject(raw_value: object) -> Optional[str]:
    """
    Kiểm tra một ô dữ liệu có phải là Mã môn học hợp lệ hay chỉ là ghi chú sự kiện.
    Trả về mã môn học đã chuẩn hóa (in hoa, không khoảng trắng thừa) nếu hợp lệ,
    ngược lại trả về None.
    """
    if raw_value is None:
        return None

    text = re.sub(r"\s+", " ", str(raw_value)).strip()
    if not text:
        return None

    # Loại bỏ các chuỗi chứa dấu tiếng Việt (ví dụ: "Nghỉ lễ 2/9", "Khai giảng K22")
    if VIETNAMESE_DIACRITICS_RE.search(text):
        return None

    # Loại bỏ các chuỗi chứa dấu gạch chéo ngày tháng hoặc câu dài
    if "/" in text or len(text) > 20:
        return None

    # Loại bỏ các từ khóa ghi chú không phải mã môn
    upper_tokens = set(re.split(r"[\s_-]+", text.upper()))
    if upper_tokens & NON_SUBJECT_KEYWORDS:
        return None

    # Chuẩn hóa bỏ khoảng trắng thừa để kiểm tra định dạng mã môn (ví dụ: "MAT 001" -> "MAT001")
    compact = re.sub(r"\s+", "", text).upper()

    # Mã môn học chuẩn: bắt đầu bằng 2-8 chữ cái, theo sau là chữ số/ký tự mã (có thể có '-' hoặc '_')
    # Yêu cầu có ít nhất 1 chữ số để phân biệt với từ thông thường như "Holiday", "Meeting"
    if re.match(r"^[A-Z]{2,8}[-_]?[A-Z0-9]*\d[A-Z0-9]*$", compact):
        return compact

    return None


def _detect_block_day_columns(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    block_row: int,
    weekday_row: int,
    block_header_text: str,
) -> int:
    """
    Xác định số cột ngày hợp lệ (bắt đầu từ cột B = 2) trong một khối lịch.
    Kết hợp kiểm tra dòng nhãn Thứ và nhãn khối Tuần (ví dụ 'Tuần 16-17' -> 2 tuần * 7 ngày = 14 cột).
    """
    # Cách 1: Đếm số cột liên tiếp có nhãn Thứ/CN ở dòng weekday_row
    count_from_weekdays = 0
    for col in range(2, ws.max_column + 1):
        val = str(ws.cell(row=weekday_row, column=col).value or "").strip()
        if val and WEEKDAY_MARKER_RE.search(val):
            count_from_weekdays += 1
        elif count_from_weekdays > 0:
            # Nếu đã bắt đầu đếm thứ mà gặp ô trống liên tiếp -> kiểm tra thêm 1 ô phòng trường hợp CN bị gộp
            next_val = (
                str(ws.cell(row=weekday_row, column=col + 1).value or "").strip()
                if col + 1 <= ws.max_column
                else ""
            )
            if not next_val:
                break

    if count_from_weekdays >= 7:
        return count_from_weekdays

    # Cách 2: Suy ra từ nhãn khối tuần (ví dụ: "Tuần 1-3" -> 3 tuần * 7 = 21 ngày)
    m = WEEK_BLOCK_HEADER_RE.match(block_header_text)
    if m:
        start_w = int(m.group(1))
        end_w = int(m.group(2)) if m.group(2) else start_w
        if end_w >= start_w:
            return (end_w - start_w + 1) * 7

    # Fallback: đếm số cột có dữ liệu ngày ở dòng weekday_row + 1
    date_row = weekday_row + 1
    count_from_dates = 0
    for col in range(2, ws.max_column + 1):
        if ws.cell(row=date_row, column=col).value is not None:
            count_from_dates += 1
        elif count_from_dates > 0:
            break
    return count_from_dates


def scan_schedule_workbook(file_path: str) -> ParseResult:
    """
    Thực hiện quét động toàn bộ file Schedule.xlsx từ trên xuống dưới.
    Hỗ trợ số lượng tuần bất kỳ (15 tuần, 17 tuần, 21 tuần...).
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Không tìm thấy file: {file_path}")

    wb = openpyxl.load_workbook(file_path, data_only=True)
    try:
        ws = wb.active
        metadata = parse_header_metadata(ws)

        records: List[SessionRecord] = []
        ignored_notes: List[Tuple[str, str, int]] = []
        detected_slots: set[int] = set()
        current_date = metadata.date_begin
        total_days_scanned = 0
        total_blocks = 0

        r = 1
        max_row = ws.max_row
        while r <= max_row:
            col_a = str(ws.cell(row=r, column=1).value or "").strip()
            block_match = WEEK_BLOCK_HEADER_RE.match(col_a)

            if not block_match:
                r += 1
                continue

            # Xác nhận đây thực sự là dòng mở đầu khối lịch (kiểm tra có dòng Tiết học phía dưới)
            # Cấu trúc chuẩn:
            #   r     : Nhãn khối tuần ("Tuần 1-3")
            #   r + 1 : Nhãn tuần con ("Tuần 1", "Tuần 2"...)
            #   r + 2 : Nhãn Thứ ("Thứ hai", "Thứ ba"... "CN")
            #   r + 3 : Dòng Ngày trong tháng
            #   r + 4..: Các dòng Tiết học ("Tiết 1".."Tiết 5" hoặc "Tiết 6".."Tiết 10")
            weekday_row = r + 2
            slot_start_row = r + 4

            # Linh hoạt tìm dòng Tiết học đầu tiên trong phạm vi r+1..r+6 (phòng trường hợp thiếu dòng tuần con)
            found_slot_start = None
            for probe in range(r + 1, min(r + 7, max_row + 1)):
                probe_a = str(ws.cell(row=probe, column=1).value or "").strip()
                if SLOT_LABEL_RE.match(probe_a):
                    found_slot_start = probe
                    break

            if found_slot_start is None:
                r += 1
                continue

            slot_start_row = found_slot_start
            # Dòng Thứ nằm ngay phía trên dòng Ngày, tức thường là slot_start_row - 2
            weekday_row = max(r + 1, slot_start_row - 2)

            # Thu thập toàn bộ các dòng Tiết học liên tiếp của khối này
            slot_rows: List[Tuple[int, int]] = []  # (row_index, slot_number)
            scan_r = slot_start_row
            while scan_r <= max_row:
                label_a = str(ws.cell(row=scan_r, column=1).value or "").strip()
                sm = SLOT_LABEL_RE.match(label_a)
                if not sm:
                    break
                slot_num = int(sm.group(1))
                slot_rows.append((scan_r, slot_num))
                detected_slots.add(slot_num)
                scan_r += 1

            if not slot_rows:
                r += 1
                continue

            # Xác định số cột ngày trong khối này
            num_day_cols = _detect_block_day_columns(ws, r, weekday_row, col_a)
            if num_day_cols <= 0:
                r = scan_r
                continue

            total_blocks += 1

            # Duyệt từng cột ngày từ trái sang phải (từ cột B = 2)
            for offset in range(num_day_cols):
                col_idx = 2 + offset
                cell_date = current_date
                date_str = cell_date.strftime("%d/%m/%Y")

                # Duyệt từng tiết học trong ngày
                for row_idx, slot_num in slot_rows:
                    raw_cell = ws.cell(row=row_idx, column=col_idx).value
                    if raw_cell is None:
                        continue
                    raw_text = str(raw_cell).strip()
                    if not raw_text:
                        continue

                    subject_code = normalize_and_validate_subject(raw_text)
                    if subject_code is None:
                        ignored_notes.append((raw_text, date_str, slot_num))
                        continue

                    records.append(
                        SessionRecord(
                            term=metadata.term,
                            date_obj=cell_date,
                            date_str=date_str,
                            slot=slot_num,
                            room=metadata.room,
                            subject=subject_code,
                            class_name=metadata.class_name,
                        )
                    )

                current_date += timedelta(days=1)
                total_days_scanned += 1

            # Nhảy qua toàn bộ khối đã quét
            r = scan_r

        if total_blocks == 0 or not records:
            raise ValueError(
                "Không quét được khối thời khóa biểu hoặc môn học hợp lệ nào trong file. "
                "Vui lòng kiểm tra lại cấu trúc file Schedule.xlsx."
            )

        # Nhận diện Ca học dựa trên danh sách tiết học
        if detected_slots:
            min_slot, max_slot = min(detected_slots), max(detected_slots)
            if min_slot >= 6:
                metadata.shift_label = f"Ca Chiều (Tiết {min_slot} - {max_slot})"
            elif max_slot <= 5:
                metadata.shift_label = f"Ca Sáng (Tiết {min_slot} - {max_slot})"
            else:
                metadata.shift_label = f"Cả ngày (Tiết {min_slot} - {max_slot})"

        metadata.total_weeks = max(1, round(total_days_scanned / 7))
        metadata.total_blocks = total_blocks

        unique_subjects = sorted({rec.subject for rec in records})
        return ParseResult(
            metadata=metadata,
            records=records,
            subjects=unique_subjects,
            ignored_notes=ignored_notes,
        )
    finally:
        wb.close()


# ==============================================================================
# XUẤT FILE EXCEL CON CHO TỪNG MÔN HỌC (EXPORTER)
# ==============================================================================

EXPORT_COLUMNS = [
    "Term",
    "Date",
    "Slot",
    "Room",
    "Lecturer",
    "Subject",
    "Session No.",
    "Class",
]


def export_subject_schedules(
    parse_result: ParseResult,
    lecturer_map: Dict[str, str],
    output_dir: str,
) -> List[str]:
    """
    Nhóm dữ liệu theo từng môn học, sắp xếp tăng dần theo thời gian,
    đánh số thứ tự buổi học (Session No.) và xuất ra các file <Class>_<Subject>.xlsx.
    Trả về danh sách đường dẫn các file đã tạo.
    """
    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    # Nhóm bản ghi theo từng môn học
    grouped: Dict[str, List[SessionRecord]] = {}
    for rec in parse_result.records:
        grouped.setdefault(rec.subject, []).append(rec)

    created_files: List[str] = []

    # Định nghĩa style openpyxl chuẩn đẹp
    header_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="0F172A")
    data_font = Font(name="Calibri", size=11, bold=False, color="1E293B")
    center_align = Alignment(horizontal="center", vertical="center")
    thin_side = Side(style="thin", color="CBD5E1")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    class_name = parse_result.metadata.class_name

    for subject in sorted(grouped.keys()):
        # Sắp xếp buổi học theo trình tự thời gian tăng dần (Ngày -> Tiết)
        sessions = sorted(grouped[subject], key=lambda item: (item.date_obj, item.slot))
        lecturer_name = lecturer_map.get(subject, "").strip()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = subject[:31]  # Tên sheet tối đa 31 ký tự trong Excel

        # Ghi dòng tiêu đề
        ws.append(EXPORT_COLUMNS)
        ws.row_dimensions[1].height = 28
        for col_idx in range(1, len(EXPORT_COLUMNS) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = cell_border

        # Ghi dữ liệu từng buổi học với Session No. tăng dần từ 1
        for idx, session in enumerate(sessions, start=1):
            row_values = [
                session.term,
                session.date_str,
                session.slot,
                session.room,
                lecturer_name,
                session.subject,
                idx,
                session.class_name,
            ]
            ws.append(row_values)
            excel_row = idx + 1
            ws.row_dimensions[excel_row].height = 22

            for col_idx in range(1, len(EXPORT_COLUMNS) + 1):
                cell = ws.cell(row=excel_row, column=col_idx)
                cell.font = data_font
                cell.alignment = center_align
                cell.border = cell_border
                if idx % 2 == 0:
                    cell.fill = zebra_fill

        # Cố định dòng tiêu đề khi cuộn
        ws.freeze_panes = "A2"

        # Tự động căn chỉnh độ rộng cột (auto-adjust column width)
        for col_idx in range(1, len(EXPORT_COLUMNS) + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for row_idx in range(1, ws.max_row + 1):
                val = ws.cell(row=row_idx, column=col_idx).value
                text_len = len(str(val)) if val is not None else 0
                if text_len > max_len:
                    max_len = text_len
            ws.column_dimensions[col_letter].width = max(max_len + 6, 12)

        filename = f"{class_name}_{subject}.xlsx"
        target_file = out_path / filename
        wb.save(str(target_file))
        wb.close()
        created_files.append(str(target_file))

    return created_files


# ==============================================================================
# GIAO DIỆN ĐỒ HỌA TKINTER (GUI)
# ==============================================================================

class ScheduleSplitterApp:
    """Ứng dụng Desktop GUI sử dụng Tkinter để bóc tách thời khóa biểu."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Công Cụ Bóc Tách Thời Khóa Biểu Theo Môn Học")
        self.root.geometry("980x720")
        self.root.minsize(860, 640)
        self.root.configure(bg="#F8FAFC")

        self.selected_file: Optional[str] = None
        self.output_dir: Optional[str] = None
        self.parse_result: Optional[ParseResult] = None
        self.lecturer_map: Dict[str, str] = {}
        self.session_counts: Dict[str, int] = {}

        self._setup_styles()
        self._build_ui()

    def _setup_styles(self) -> None:
        """Thiết lập giao diện ttk hiện đại, sắc nét."""
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("App.TFrame", background="#F8FAFC")
        style.configure("Card.TFrame", background="#FFFFFF", relief="flat")
        style.configure(
            "Title.TLabel",
            background="#F8FAFC",
            foreground="#0F172A",
            font=("Segoe UI", 16, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background="#F8FAFC",
            foreground="#475569",
            font=("Segoe UI", 10),
        )
        style.configure(
            "Section.TLabel",
            background="#FFFFFF",
            foreground="#0F172A",
            font=("Segoe UI", 11, "bold"),
        )
        style.configure(
            "MetaKey.TLabel",
            background="#FFFFFF",
            foreground="#64748B",
            font=("Segoe UI", 9),
        )
        style.configure(
            "MetaVal.TLabel",
            background="#FFFFFF",
            foreground="#0F172A",
            font=("Segoe UI", 10, "bold"),
        )
        style.configure(
            "Primary.TButton",
            background="#1D4ED8",
            foreground="#FFFFFF",
            font=("Segoe UI", 10, "bold"),
            padding=(16, 8),
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#1E40AF"), ("disabled", "#94A3B8")],
            foreground=[("disabled", "#F1F5F9")],
        )
        style.configure(
            "Secondary.TButton",
            background="#E2E8F0",
            foreground="#0F172A",
            font=("Segoe UI", 9, "bold"),
            padding=(12, 6),
            borderwidth=0,
        )
        style.map("Secondary.TButton", background=[("active", "#CBD5E1")])

        style.configure(
            "Treeview",
            background="#FFFFFF",
            foreground="#0F172A",
            fieldbackground="#FFFFFF",
            rowheight=30,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Treeview.Heading",
            background="#E2E8F0",
            foreground="#0F172A",
            font=("Segoe UI", 10, "bold"),
            padding=(8, 6),
        )
        style.map("Treeview", background=[("selected", "#DBEAFE")], foreground=[("selected", "#1E3A8A")])

    def _build_ui(self) -> None:
        """Xây dựng bố cục giao diện người dùng."""
        container = ttk.Frame(self.root, style="App.TFrame", padding=20)
        container.pack(fill=tk.BOTH, expand=True)

        # --- Header ---
        header_frame = ttk.Frame(container, style="App.TFrame")
        header_frame.pack(fill=tk.X, pady=(0, 14))

        ttk.Label(
            header_frame,
            text="Tự Động Bóc Tách Thời Khóa Biểu (Dynamic Schedule Splitter)",
            style="Title.TLabel",
        ).pack(anchor=tk.W)
        ttk.Label(
            header_frame,
            text="Quét động toàn bộ các khối tuần (15, 17, 21 tuần hoặc bất kỳ), gán giảng viên và xuất file Excel chuẩn theo từng môn.",
            style="Subtitle.TLabel",
        ).pack(anchor=tk.W, pady=(2, 0))

        # --- Khối 1: Chọn file nguồn ---
        file_card = tk.Frame(container, bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1, padx=16, pady=12)
        file_card.pack(fill=tk.X, pady=(0, 12))

        file_row = tk.Frame(file_card, bg="#FFFFFF")
        file_row.pack(fill=tk.X)

        self.btn_pick_file = ttk.Button(
            file_row,
            text="Chọn file Schedule.xlsx",
            style="Primary.TButton",
            command=self.on_pick_file,
        )
        self.btn_pick_file.pack(side=tk.LEFT)

        self.lbl_file_path = tk.Label(
            file_row,
            text="Chưa chọn file nào...",
            bg="#F1F5F9",
            fg="#475569",
            font=("Consolas", 9),
            anchor=tk.W,
            padx=10,
            pady=7,
        )
        self.lbl_file_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(12, 0))

        # --- Khối 2: Thông tin tự động bóc tách ---
        meta_card = tk.Frame(container, bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1, padx=16, pady=12)
        meta_card.pack(fill=tk.X, pady=(0, 12))

        ttk.Label(meta_card, text="Thông Tin Bóc Tách Tự Động", style="Section.TLabel").pack(anchor=tk.W, pady=(0, 8))

        meta_grid = tk.Frame(meta_card, bg="#FFFFFF")
        meta_grid.pack(fill=tk.X)
        for i in range(3):
            meta_grid.columnconfigure(i, weight=1)

        self.meta_vars: Dict[str, tk.StringVar] = {
            "class": tk.StringVar(value="—"),
            "room": tk.StringVar(value="—"),
            "term": tk.StringVar(value="—"),
            "daterange": tk.StringVar(value="—"),
            "shift": tk.StringVar(value="—"),
            "weeks": tk.StringVar(value="—"),
        }

        meta_items = [
            ("Lớp học (Class)", "class", 0, 0),
            ("Phòng học (Room)", "room", 0, 1),
            ("Học kỳ (Term)", "term", 0, 2),
            ("Khoảng ngày học", "daterange", 1, 0),
            ("Ca học nhận diện", "shift", 1, 1),
            ("Tổng số tuần & buổi", "weeks", 1, 2),
        ]

        for label_text, key, r_idx, c_idx in meta_items:
            cell = tk.Frame(meta_grid, bg="#FFFFFF", pady=4)
            cell.grid(row=r_idx, column=c_idx, sticky=tk.W, padx=(0, 16))
            ttk.Label(cell, text=label_text, style="MetaKey.TLabel").pack(anchor=tk.W)
            ttk.Label(cell, textvariable=self.meta_vars[key], style="MetaVal.TLabel").pack(anchor=tk.W)

        # --- Khối 3: Danh sách môn học & Phân công giảng viên ---
        subj_card = tk.Frame(container, bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1, padx=16, pady=12)
        subj_card.pack(fill=tk.BOTH, expand=True, pady=(0, 12))

        subj_header = tk.Frame(subj_card, bg="#FFFFFF")
        subj_header.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(subj_header, text="Danh Sách Môn Học & Phân Công Giảng Viên", style="Section.TLabel").pack(side=tk.LEFT)

        self.lbl_ignored_info = tk.Label(
            subj_header,
            text="",
            bg="#FFFFFF",
            fg="#64748B",
            font=("Segoe UI", 9),
        )
        self.lbl_ignored_info.pack(side=tk.RIGHT)

        # Bảng Treeview
        tree_frame = tk.Frame(subj_card, bg="#FFFFFF")
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("subject", "sessions", "first_date", "last_date", "lecturer")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("subject", text="Mã môn học (Subject)")
        self.tree.heading("sessions", text="Số buổi")
        self.tree.heading("first_date", text="Buổi đầu")
        self.tree.heading("last_date", text="Buổi cuối")
        self.tree.heading("lecturer", text="Tên Giảng viên (Lecturer)")

        self.tree.column("subject", width=150, anchor=tk.CENTER)
        self.tree.column("sessions", width=90, anchor=tk.CENTER)
        self.tree.column("first_date", width=120, anchor=tk.CENTER)
        self.tree.column("last_date", width=120, anchor=tk.CENTER)
        self.tree.column("lecturer", width=320, anchor=tk.W)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.bind("<Double-1>", self.on_tree_double_click)

        # Thanh nhập nhanh tên giảng viên cho môn đang chọn
        edit_bar = tk.Frame(subj_card, bg="#F8FAFC", highlightbackground="#E2E8F0", highlightthickness=1, padx=10, pady=8)
        edit_bar.pack(fill=tk.X, pady=(10, 0))

        tk.Label(edit_bar, text="Môn đang chọn:", bg="#F8FAFC", fg="#475569", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self.var_active_subject = tk.StringVar(value="—")
        tk.Label(
            edit_bar,
            textvariable=self.var_active_subject,
            bg="#F8FAFC",
            fg="#1D4ED8",
            font=("Segoe UI", 10, "bold"),
            width=12,
            anchor=tk.W,
        ).pack(side=tk.LEFT, padx=(4, 12))

        tk.Label(edit_bar, text="Tên Giảng viên:", bg="#F8FAFC", fg="#475569", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self.var_lecturer_input = tk.StringVar()
        self.entry_lecturer = ttk.Entry(edit_bar, textvariable=self.var_lecturer_input, font=("Segoe UI", 10), width=34)
        self.entry_lecturer.pack(side=tk.LEFT, padx=(6, 8), fill=tk.X, expand=True)
        self.entry_lecturer.bind("<Return>", lambda _e: self.on_apply_lecturer(move_next=True))

        self.btn_apply_lecturer = ttk.Button(
            edit_bar,
            text="Cập nhật giảng viên",
            style="Secondary.TButton",
            command=lambda: self.on_apply_lecturer(move_next=True),
            state=tk.DISABLED,
        )
        self.btn_apply_lecturer.pack(side=tk.LEFT)

        # --- Khối 4: Chọn thư mục lưu & Xuất file ---
        action_bar = tk.Frame(container, bg="#F8FAFC")
        action_bar.pack(fill=tk.X)

        self.btn_pick_out = ttk.Button(
            action_bar,
            text="Chọn thư mục lưu...",
            style="Secondary.TButton",
            command=self.on_pick_output_dir,
        )
        self.btn_pick_out.pack(side=tk.LEFT)

        self.lbl_output_dir = tk.Label(
            action_bar,
            text="Thư mục xuất: (Mặc định cùng thư mục với file nguồn)",
            bg="#F8FAFC",
            fg="#475569",
            font=("Segoe UI", 9),
            anchor=tk.W,
        )
        self.lbl_output_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 12))

        self.btn_export = ttk.Button(
            action_bar,
            text="Bắt đầu xuất file",
            style="Primary.TButton",
            command=self.on_export_files,
            state=tk.DISABLED,
        )
        self.btn_export.pack(side=tk.RIGHT)

    # --------------------------------------------------------------------------
    # CÁC HÀM XỬ LÝ SỰ KIỆN GIAO DIỆN (EVENT HANDLERS)
    # --------------------------------------------------------------------------

    def on_pick_file(self) -> None:
        """Mở hộp thoại chọn file Schedule.xlsx và phân tích tự động."""
        file_path = filedialog.askopenfilename(
            title="Chọn file thời khóa biểu (Schedule.xlsx)",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Tất cả file", "*.*")],
        )
        if not file_path:
            return

        try:
            result = scan_schedule_workbook(file_path)
        except Exception as exc:
            messagebox.showerror("Lỗi đọc thời khóa biểu", f"Không thể xử lý file:\n{exc}")
            return

        self.selected_file = file_path
        self.parse_result = result
        self.lbl_file_path.config(text=file_path, fg="#0F172A")

        # Mặc định thư mục xuất là thư mục chứa file nguồn / output_schedules
        default_out = str(Path(file_path).parent / f"Output_{result.metadata.class_name}")
        self.output_dir = default_out
        self.lbl_output_dir.config(text=f"Thư mục xuất: {default_out}")

        # Cập nhật thông tin metadata lên giao diện
        meta = result.metadata
        self.meta_vars["class"].set(meta.class_name)
        self.meta_vars["room"].set(meta.room)
        self.meta_vars["term"].set(meta.term)
        self.meta_vars["daterange"].set(
            f"{meta.date_begin.strftime('%d/%m/%Y')} - {meta.date_end.strftime('%d/%m/%Y')}"
        )
        self.meta_vars["shift"].set(meta.shift_label)
        self.meta_vars["weeks"].set(
            f"{meta.total_weeks} tuần ({meta.total_blocks} khối) · {len(result.records)} buổi học"
        )

        if result.ignored_notes:
            self.lbl_ignored_info.config(
                text=f"Đã bỏ qua {len(result.ignored_notes)} ô ghi chú/ngày nghỉ"
            )
        else:
            self.lbl_ignored_info.config(text="")

        # Nạp danh sách môn học vào bảng Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.lecturer_map = {subj: "" for subj in result.subjects}
        grouped: Dict[str, List[SessionRecord]] = {}
        for rec in result.records:
            grouped.setdefault(rec.subject, []).append(rec)

        for subj in result.subjects:
            sessions = sorted(grouped[subj], key=lambda s: (s.date_obj, s.slot))
            first_d = sessions[0].date_str
            last_d = sessions[-1].date_str
            self.tree.insert(
                "",
                tk.END,
                iid=subj,
                values=(subj, len(sessions), first_d, last_d, ""),
            )

        # Tự động chọn môn đầu tiên để người dùng nhập nhanh giảng viên
        if result.subjects:
            first_subj = result.subjects[0]
            self.tree.selection_set(first_subj)
            self.tree.focus(first_subj)
            self.var_active_subject.set(first_subj)
            self.var_lecturer_input.set("")
            self.btn_apply_lecturer.config(state=tk.NORMAL)
            self.btn_export.config(state=tk.NORMAL)
            self.entry_lecturer.focus_set()

    def on_tree_select(self, _event: tk.Event) -> None:
        """Đồng bộ môn học được chọn xuống ô nhập Giảng viên."""
        selected = self.tree.selection()
        if not selected:
            return
        subj = selected[0]
        self.var_active_subject.set(subj)
        self.var_lecturer_input.set(self.lecturer_map.get(subj, ""))

    def on_tree_double_click(self, _event: tk.Event) -> None:
        """Khi nhấp đúp vào một dòng môn học, chuyển con trỏ vào ô nhập Giảng viên."""
        self.on_tree_select(_event)
        self.entry_lecturer.focus_set()
        self.entry_lecturer.select_range(0, tk.END)

    def on_apply_lecturer(self, move_next: bool = False) -> None:
        """Lưu tên Giảng viên cho môn học đang chọn và chuyển sang môn tiếp theo nếu cần."""
        selected = self.tree.selection()
        if not selected:
            return
        subj = selected[0]
        lecturer_name = self.var_lecturer_input.get().strip()
        self.lecturer_map[subj] = lecturer_name

        current_vals = list(self.tree.item(subj, "values"))
        current_vals[4] = lecturer_name
        self.tree.item(subj, values=current_vals)

        if move_next:
            items = self.tree.get_children()
            idx = items.index(subj)
            if idx + 1 < len(items):
                next_subj = items[idx + 1]
                self.tree.selection_set(next_subj)
                self.tree.focus(next_subj)
                self.tree.see(next_subj)
                self.var_active_subject.set(next_subj)
                self.var_lecturer_input.set(self.lecturer_map.get(next_subj, ""))
                self.entry_lecturer.focus_set()

    def on_pick_output_dir(self) -> None:
        """Cho phép người dùng chọn thư mục lưu các file Excel kết quả."""
        chosen = filedialog.askdirectory(title="Chọn thư mục lưu các file lịch môn học")
        if chosen:
            self.output_dir = chosen
            self.lbl_output_dir.config(text=f"Thư mục xuất: {chosen}")

    def on_export_files(self) -> None:
        """Thực hiện xuất toàn bộ các file <Class>_<Subject>.xlsx."""
        if not self.parse_result:
            messagebox.showwarning("Chưa có dữ liệu", "Vui lòng chọn file Schedule.xlsx trước.")
            return

        # Lưu lại giá trị đang gõ dở trong ô nhập Giảng viên nếu có
        selected = self.tree.selection()
        if selected:
            subj = selected[0]
            typed = self.var_lecturer_input.get().strip()
            if typed != self.lecturer_map.get(subj, ""):
                self.lecturer_map[subj] = typed
                vals = list(self.tree.item(subj, "values"))
                vals[4] = typed
                self.tree.item(subj, values=vals)

        # Kiểm tra xem có môn nào chưa nhập Giảng viên không (cảnh báo nhẹ, vẫn cho phép xuất)
        missing_lecturers = [s for s in self.parse_result.subjects if not self.lecturer_map.get(s, "").strip()]
        if missing_lecturers:
            proceed = messagebox.askyesno(
                "Thiếu tên Giảng viên",
                f"Các môn sau chưa có tên Giảng viên:\n{', '.join(missing_lecturers)}\n\n"
                "Bạn vẫn muốn tiếp tục xuất file (cột Lecturer sẽ để trống)?",
            )
            if not proceed:
                return

        target_dir = self.output_dir or str(Path(self.selected_file).parent)
        try:
            created_files = export_subject_schedules(
                parse_result=self.parse_result,
                lecturer_map=self.lecturer_map,
                output_dir=target_dir,
            )
        except Exception as exc:
            messagebox.showerror("Lỗi xuất file", f"Quá trình xuất file gặp lỗi:\n{exc}")
            return

        file_names = "\n".join(f"• {Path(p).name}" for p in created_files)
        messagebox.showinfo(
            "Xuất file hoàn tất",
            f"Đã tạo thành công {len(created_files)} file Excel theo từng môn học!\n\n"
            f"Thư mục lưu:\n{target_dir}\n\n"
            f"Danh sách file:\n{file_names}",
        )


def main() -> None:
    root = tk.Tk()
    ScheduleSplitterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
