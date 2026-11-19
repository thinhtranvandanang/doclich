# Công cụ Tự động Bóc tách Thời khóa biểu (Dynamic Schedule Splitter)

Ứng dụng Python Desktop (Tkinter + OpenPyXL) giúp đọc file thời khóa biểu tổng quát (`Schedule.xlsx`) với **cơ chế quét động (Dynamic Block Scanning)** cho mọi số lượng tuần (15, 17, 21 tuần hoặc bất kỳ), phân công Giảng viên và xuất tự động từng file `<Class>_<Subject>.xlsx`.

## Cài đặt & Chạy

```bash
pip install -r requirements.txt
python schedule_splitter.py
```

### Tạo file mẫu để kiểm thử (17 tuần & 21 tuần)

```bash
python generate_sample_schedule.py
```
