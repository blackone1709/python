import tkinter as tk
from tkinter import messagebox
from tkcalendar import Calendar
import csv
from datetime import datetime
import os
import pandas as pd
from tkinter.filedialog import asksaveasfilename

root = tk.Tk()
root.title("Save Data with Options and Age")
root.geometry("600x700")

CSV_FILE = "thongtin.csv"

fields = ["Mã", "Tên", "Đơn vị", "Chức danh", "Số CMND", "Nơi cấp"]
entries = {}


def create_label_entry(parent, text, row):
    label = tk.Label(parent, text=text, anchor='w')
    label.grid(row=row, column=0, sticky='w', padx=10, pady=5)
    entry = tk.Entry(parent)
    entry.grid(row=row, column=1, padx=10, pady=5)
    return entry


input_frame = tk.LabelFrame(root, text="Nhập thông tin nhân viên")
input_frame.pack(fill="x", padx=10, pady=10)

for i, field in enumerate(fields):
    entries[field] = create_label_entry(input_frame, f"{field}:", i)

birthday_label = tk.Label(input_frame, text="Sinh nhật:")
birthday_label.grid(row=len(fields), column=0, sticky='w', padx=10, pady=5)
birthday_calendar = Calendar(input_frame, selectmode='day', year=2000, month=1, day=1)
birthday_calendar.grid(row=len(fields), column=1, padx=10, pady=5)

option_frame = tk.LabelFrame(root, text="Tùy chọn")
option_frame.pack(fill="x", padx=10, pady=10)

type_label = tk.Label(option_frame, text="Loại:")
type_label.grid(row=0, column=0, sticky='w', padx=10, pady=5)

type_var = tk.IntVar(value=0)
tk.Radiobutton(option_frame, text="Khách hàng", variable=type_var, value=1).grid(row=0, column=1, padx=5)
tk.Radiobutton(option_frame, text="Cung cấp", variable=type_var, value=2).grid(row=0, column=2, padx=5)

gender_label = tk.Label(option_frame, text="Giới tính:")
gender_label.grid(row=1, column=0, sticky='w', padx=10, pady=5)

gender_var = tk.IntVar(value=0)
tk.Radiobutton(option_frame, text="Nam", variable=gender_var, value=3).grid(row=1, column=1, padx=5)
tk.Radiobutton(option_frame, text="Nữ", variable=gender_var, value=4).grid(row=1, column=2, padx=5)


def calculate_age(birthday_str):
    try:
        birth_date = datetime.strptime(birthday_str, "%m/%d/%y")  # Định dạng ngày mm/dd/yy
        today = datetime.today()
        age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
        return age
    except ValueError as e:
        raise ValueError("Sai định dạng ngày tháng. Hãy chọn ngày hợp lệ từ lịch.") from e


def save_to_csv():
    user_data = {}
    for field, entry in entries.items():
        value = entry.get().strip()
        if not value:
            messagebox.showerror("Input Error", f"Trường '{field}' không được bỏ trống!")
            return
        user_data[field] = value

    birthday = birthday_calendar.get_date()
    user_data["Sinh"] = birthday

    try:
        age = calculate_age(birthday)
        user_data["Age"] = age
    except Exception as e:
        messagebox.showerror("Error", f"Lỗi tính tuổi: {e}")
        return

    type_value = type_var.get()
    gender_value = gender_var.get()

    if type_value == 0:
        messagebox.showerror("Input Error", "Vui lòng chọn Loại (Khách hàng hoặc Cung cấp)!")
        return
    if gender_value == 0:
        messagebox.showerror("Input Error", "Vui lòng chọn Giới tính (Nam hoặc Nữ)!")
        return

    type_mapping = {1: "Khách hàng", 2: "Cung cấp"}
    gender_mapping = {3: "Nam", 4: "Nữ"}
    user_data["Type"] = type_mapping[type_value]
    user_data["Gender"] = gender_mapping[gender_value]

    try:
        file_exists = os.path.exists(CSV_FILE)
        with open(CSV_FILE, mode='a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            if not file_exists:
                writer.writerow([*fields, "Sinh", "Age", "Type", "Gender"])
            writer.writerow([user_data.get(field, '') for field in [*fields, "Sinh", "Age", "Type", "Gender"]])

        messagebox.showinfo("Success", "Dữ liệu đã được lưu thành công!")
        for entry in entries.values():
            entry.delete(0, tk.END)
        birthday_calendar.selection_set("2000-01-01")
        type_var.set(0)
        gender_var.set(0)
    except Exception as e:
        messagebox.showerror("Error", f"Lỗi khi lưu dữ liệu: {e}")


def show_today_birthdays():
    if not os.path.exists(CSV_FILE):
        messagebox.showinfo("Thông báo", "Không có dữ liệu để hiển thị.")
        return

    today = datetime.today().strftime("%m/%d/%y")
    with open(CSV_FILE, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        today_birthdays = [row for row in reader if row.get("Sinh") == today]

    if today_birthdays:
        message = "Danh sách nhân viên sinh nhật hôm nay:\n"
        for person in today_birthdays:
            message += f"- {person['Tên']} ({person['Sinh']})\n"
        messagebox.showinfo("Sinh nhật hôm nay", message)
    else:
        messagebox.showinfo("Sinh nhật hôm nay", "Hôm nay không có nhân viên nào sinh nhật.")


def export_sorted_by_age():
    try:
        if not os.path.exists(CSV_FILE):
            messagebox.showerror("Lỗi", "Không tìm thấy tệp dữ liệu để xuất!")
            return

        data = pd.read_csv(CSV_FILE)

        data_sorted = data.sort_values(by="Age", ascending=False)

        file_path = asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            title="Lưu file Excel",
            initialfile="DanhSachNhanVien.xlsx"
        )

        if not file_path:
            messagebox.showinfo("Hủy bỏ", "Bạn đã hủy lưu file Excel.")
            return

        data_sorted.to_excel(file_path, index=False, engine='openpyxl')

        messagebox.showinfo("Thành công", f"File Excel đã được lưu tại:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Đã xảy ra lỗi khi xuất file Excel: {e}")


button_frame = tk.Frame(root)
button_frame.pack(pady=10)

tk.Button(button_frame, text="Lưu", width=15, command=save_to_csv).grid(row=0, column=0, padx=10)
tk.Button(button_frame, text="Sinh nhật hôm nay", width=20, command=show_today_birthdays).grid(row=0, column=1, padx=10)
tk.Button(button_frame, text="Xuất danh sách", width=20, command=export_sorted_by_age).grid(row=0, column=2, padx=10)

root.mainloop()
