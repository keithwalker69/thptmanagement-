import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, Listbox, Toplevel
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import os
import numpy as np
import glob

# --- CÁC HÀM PHÂN TÍCH VÀ XỬ LÝ DỮ LIỆU ---

CACHE_DIR = 'data_cache'

PROVINCE_CODES = {
    '01': 'Hà Nội', '02': 'TP. HCM', '03': 'Hải Phòng', '04': 'Đà Nẵng',
    '05': 'Hà Giang', '06': 'Cao Bằng', '07': 'Lai Châu', '08': 'Lào Cai',
    '09': 'Tuyên Quang', '10': 'Lạng Sơn', '11': 'Bắc Kạn', '12': 'Thái Nguyên',
    '13': 'Yên Bái', '14': 'Sơn La', '15': 'Phú Thọ', '16': 'Vĩnh Phúc',
    '17': 'Quảng Ninh', '18': 'Bắc Giang', '19': 'Bắc Ninh', '21': 'Hải Dương',
    '22': 'Hưng Yên', '23': 'Hòa Bình', '24': 'Hà Nam', '25': 'Nam Định',
    '26': 'Thái Bình', '27': 'Ninh Bình', '28': 'Thanh Hóa', '29': 'Nghệ An',
    '30': 'Hà Tĩnh', '31': 'Quảng Bình', '32': 'Quảng Trị', '33': 'Thừa Thiên Huế',
    '34': 'Quảng Nam', '35': 'Quảng Ngãi', '36': 'Kon Tum', '37': 'Bình Định',
    '38': 'Gia Lai', '39': 'Phú Yên', '40': 'Đắk Lắk', '41': 'Khánh Hòa',
    '42': 'Lâm Đồng', '43': 'Bình Phước', '44': 'Bình Dương', '45': 'Ninh Thuận',
    '46': 'Tây Ninh', '47': 'Bình Thuận', '48': 'Đồng Nai', '49': 'Long An',
    '50': 'Đồng Tháp', '51': 'An Giang', '52': 'Bà Rịa - Vũng Tàu', '53': 'Tiền Giang',
    '54': 'Kiên Giang', '55': 'Cần Thơ', '56': 'Bến Tre', '57': 'Vĩnh Long',
    '58': 'Trà Vinh', '59': 'Sóc Trăng', '60': 'Bạc Liêu', '61': 'Cà Mau',
    '62': 'Điện Biên', '63': 'Đăk Nông', '64': 'Hậu Giang'
}

def load_and_process_from_excel():
    search_pattern = '*-ketquathi-ct*.xlsx'
    excel_files = glob.glob(search_pattern)
    if not excel_files:
        messagebox.showerror("Lỗi", f"Không tìm thấy file Excel nào khớp với mẫu '{search_pattern}'.")
        return None, None
    list_of_dataframes = [pd.read_excel(file) for file in excel_files]
    df_combined = pd.concat(list_of_dataframes, ignore_index=True)
    df_combined['Mã tỉnh'] = df_combined['SOBAODANH'].astype(str).str.zfill(8).str[:2]
    df_combined['Tỉnh'] = df_combined['Mã tỉnh'].map(PROVINCE_CODES)
    loaded_files_str = "\n- ".join(excel_files)
    return df_combined, loaded_files_str

# --- CÁC HÀM VẼ BIỂU ĐỒ ---
def classify_score(score):
    if score >= 8: return 'Giỏi'
    elif score >= 6.5: return 'Khá'
    elif score >= 5: return 'Trung bình'
    else: return 'Yếu'

def plot_score_distribution(df, subject, canvas_widget):
    plt.style.use('dark_background'); fig, ax = plt.subplots(figsize=(10, 5)); scores = df[subject].dropna()
    counts, bins, patches = ax.hist(scores, bins=20, edgecolor='white', alpha=0.7, picker=True)
    ax.set_title(f'Phân phối điểm môn {subject}', fontsize=16); ax.set_xlabel('Điểm'); ax.set_ylabel('Số lượng thí sinh')
    ax.grid(axis='y', linestyle='--', alpha=0.6); ax.set_xticks(np.arange(0, 11, 1)); ax.set_xlim(-0.5, 10.5)
    def on_pick(event):
        try: idx = patches.index(event.artist)
        except ValueError: return
        student_count, score_start, score_end = int(counts[idx]), bins[idx], bins[idx+1]
        messagebox.showinfo("Thông tin chi tiết", f"Số lượng: {student_count} thí sinh\nTrong khoảng điểm: [{score_start:.2f} - {score_end:.2f}]")
    fig.canvas.mpl_connect('pick_event', on_pick); draw_figure(canvas_widget, fig)

def plot_classification_pie(df, subject, canvas_widget):
    plt.style.use('dark_background'); fig, ax = plt.subplots(figsize=(10, 5))
    data = df[subject].dropna().apply(classify_score).value_counts()
    colors = {'Giỏi': '#4CAF50', 'Khá': '#2196F3', 'Trung bình': '#FFC107', 'Yếu': '#F44336'}
    pie_colors = [colors.get(label, '#808080') for label in data.index]
    ax.pie(data, labels=data.index, autopct='%1.1f%%', startangle=140, colors=pie_colors, wedgeprops={"edgecolor":"white",'linewidth': 1, 'antialiased': True})
    ax.set_title(f'Tỷ lệ xếp loại môn {subject}', fontsize=16); draw_figure(canvas_widget, fig)

def plot_province_comparison(df, subject, canvas_widget):
    plt.style.use('dark_background'); fig, ax = plt.subplots(figsize=(10, 12))
    province_avg = df.groupby('Tỉnh')[subject].mean().dropna().sort_values(ascending=True)
    province_avg.plot(kind='barh', ax=ax, color='cyan')
    ax.set_title(f'Điểm trung bình môn {subject} theo Tỉnh/Thành phố', fontsize=16); ax.set_xlabel(f'Điểm trung bình {subject}')
    ax.set_ylabel('Tỉnh/Thành phố'); ax.grid(axis='x', linestyle='--', alpha=0.6); fig.tight_layout(); draw_figure(canvas_widget, fig)

def draw_figure(canvas, figure):
    for widget in canvas.winfo_children(): widget.destroy()
    figure_canvas = FigureCanvasTkAgg(figure, master=canvas)
    figure_canvas.draw(); figure_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

class SelectionDialog(Toplevel):
    def __init__(self, parent, file_list):
        super().__init__(parent); self.title("Chọn bộ dữ liệu"); self.geometry("300x400"); self.result = None
        self.listbox = Listbox(self, selectmode=tk.SINGLE, background="#333", foreground="white", selectbackground="#007bff")
        for f in file_list: self.listbox.insert(tk.END, os.path.basename(f))
        self.listbox.pack(pady=10, padx=10, fill=tk.BOTH, expand=True); self.listbox.bind("<Double-1>", self.on_ok)
        ok_button = tb.Button(self, text="Mở", command=self.on_ok, bootstyle=SUCCESS); ok_button.pack(pady=10)
    def on_ok(self, event=None):
        selected_indices = self.listbox.curselection()
        if selected_indices: self.result = self.listbox.get(selected_indices[0]); self.destroy()

# --- GIAO DIỆN NGƯỜI DÙNG (GUI) ---
class App(tb.Window):
    def __init__(self):
        super().__init__(themename="darkly"); self.title("Ứng dụng Phân tích Điểm thi THPT Quốc gia"); self.geometry("1200x800")
        self.df = None
        self.subjects = ['Toán', 'Văn', 'Lí', 'Hóa', 'Sinh', 'Sử', 'Địa', 'Giáo dục công dân', 'Ngoại ngữ']
        os.makedirs(CACHE_DIR, exist_ok=True)
        self.create_widgets()

    def create_widgets(self):
        main_frame = tb.Frame(self, padding=10); main_frame.pack(fill=BOTH, expand=YES)
        control_frame = tb.Labelframe(main_frame, text="Bảng điều khiển", padding=10); control_frame.pack(side=LEFT, fill=Y, padx=10, pady=10)
        self.plot_frame = tb.Frame(main_frame, padding=10); self.plot_frame.pack(side=RIGHT, fill=BOTH, expand=YES, padx=10, pady=10)

        load_button = tb.Button(control_frame, text="Mở bộ dữ liệu đã lưu", command=self.handle_load_from_cache, bootstyle=SUCCESS)
        load_button.pack(pady=(10,5), padx=10, fill=X)
        import_button = tb.Button(control_frame, text="Nhập & Lưu mới từ Excel", command=self.handle_import_and_save, bootstyle=(INFO, OUTLINE))
        import_button.pack(pady=5, padx=10, fill=X)
        self.status_label = tb.Label(control_frame, text="Chưa có dữ liệu nào được nạp", bootstyle=INVERSE)
        self.status_label.pack(pady=(10, 10), padx=10, fill=X)
        
        separator1 = tb.Separator(control_frame, orient=HORIZONTAL); separator1.pack(fill=X, pady=10, padx=5)

        subject_label = tb.Label(control_frame, text="Chọn môn học:"); subject_label.pack(pady=(10,0))
        self.subject_combo = tb.Combobox(control_frame, values=self.subjects, state="readonly"); self.subject_combo.pack(pady=5, padx=10)
        self.subject_combo.set('Toán')

        dist_button = tb.Button(control_frame, text="Phân phối điểm", command=self.handle_distribution, bootstyle=PRIMARY)
        dist_button.pack(pady=10, padx=10, fill=X)
        class_button = tb.Button(control_frame, text="Tỷ lệ Giỏi/Khá/TB/Yếu", command=self.handle_classification, bootstyle=INFO)
        class_button.pack(pady=10, padx=10, fill=X)
        province_button = tb.Button(control_frame, text="So sánh điểm các tỉnh", command=self.handle_province_comparison, bootstyle=WARNING)
        province_button.pack(pady=10, padx=10, fill=X)

        separator2 = tb.Separator(control_frame, orient=HORIZONTAL); separator2.pack(fill=X, pady=10, padx=5)
        failure_button = tb.Button(control_frame, text="Phân tích Tốt nghiệp", command=self.handle_failure_rate, bootstyle=DANGER)
        failure_button.pack(pady=10, padx=10, fill=X)

    def handle_load_from_cache(self):
        saved_files = glob.glob(os.path.join(CACHE_DIR, '*.feather'))
        if not saved_files: messagebox.showinfo("Thông báo", "Không có bộ dữ liệu nào được lưu trữ."); return
        dialog = SelectionDialog(self, saved_files); self.wait_window(dialog)
        if dialog.result:
            file_to_load = os.path.join(CACHE_DIR, dialog.result)
            try:
                self.df = pd.read_feather(file_to_load); self.status_label.config(text=f"Đang dùng: {dialog.result}")
                messagebox.showinfo("Thành công", f"Đã nạp {len(self.df)} dòng từ file '{dialog.result}'.")
            except Exception as e: messagebox.showerror("Lỗi", f"Không thể đọc file {dialog.result}.\nLỗi: {e}")

    def handle_import_and_save(self):
        try:
            new_df, loaded_files_str = load_and_process_from_excel()
            if new_df is None: return
            save_name = simpledialog.askstring("Lưu bộ dữ liệu", "Nhập tên cho bộ dữ liệu này:", parent=self)
            if save_name:
                filename = f"{save_name}.feather"; filepath = os.path.join(CACHE_DIR, filename)
                if os.path.exists(filepath) and not messagebox.askyesno("Xác nhận", f"File '{filename}' đã tồn tại. Ghi đè?"): return
                new_df.to_feather(filepath); self.df = new_df; self.status_label.config(text=f"Đang dùng: {filename}")
                messagebox.showinfo("Thành công", f"Đã xử lý và lưu thành công file '{filename}' trong thư mục {CACHE_DIR}.")
            else: messagebox.showinfo("Đã hủy", "Hành động nhập và lưu đã bị hủy.")
        except Exception as e: messagebox.showerror("Lỗi", f"Có lỗi xảy ra khi xử lý file Excel: {e}")

    # --- CẬP NHẬT LOGIC HOÀN CHỈNH ---
    def handle_failure_rate(self):
        if self.df is None:
            messagebox.showwarning("Chưa có dữ liệu", "Vui lòng 'Mở bộ dữ liệu' hoặc 'Nhập mới' trước khi phân tích.")
            return

        # Bước 1: Lọc ra những thí sinh đủ điều kiện xét tốt nghiệp (có điểm >= 4 môn)
        # Đây là bước quan trọng nhất để loại bỏ thí sinh tự do.
        eligible_mask = self.df[self.subjects].notna().sum(axis=1) >= 4
        df_eligible = self.df[eligible_mask]

        # Bước 2: Đếm số lượng các nhóm
        total_students = len(self.df)
        eligible_count = len(df_eligible)
        ineligible_count = total_students - eligible_count

        # Bước 3: Trong nhóm ĐỦ ĐIỀU KIỆN, tìm những người trượt do điểm liệt
        if eligible_count > 0:
            failed_liet_mask = (df_eligible[self.subjects] <= 1.0).any(axis=1)
            failed_liet_count = failed_liet_mask.sum()
            # Tỷ lệ trượt được tính trên số người đủ điều kiện, không phải trên tổng số
            failure_percentage = (failed_liet_count / eligible_count) * 100
        else:
            failed_liet_count = 0
            failure_percentage = 0.0

        # Bước 4: Hiển thị kết quả phân tích chi tiết
        result_message = (
            f"--- Phân tích Toàn bộ Dữ liệu ---\n"
            f"Tổng số thí sinh trong tệp: {total_students:,}\n\n"
            
            f"--- Lọc Đối tượng Xét Tốt nghiệp ---\n"
            f"Số thí sinh không xét tốt nghiệp (thi < 4 môn, thường là thí sinh tự do): {ineligible_count:,}\n"
            f"Số thí sinh đủ điều kiện xét tốt nghiệp (thi >= 4 môn): {eligible_count:,}\n\n"
            
            f"--- Kết quả Tốt nghiệp (của nhóm đủ điều kiện) ---\n"
            f"Số thí sinh trượt do có điểm liệt (<= 1.0): {failed_liet_count:,}\n\n"
            
            f"=> Tỷ lệ trượt tốt nghiệp thực tế: {failure_percentage:.2f}%"
        )
        
        messagebox.showinfo("Phân tích Kết quả Tốt nghiệp", result_message)

    def run_analysis(self, analysis_func):
        if self.df is None:
            messagebox.showwarning("Chưa có dữ liệu", "Vui lòng 'Mở bộ dữ liệu' hoặc 'Nhập mới' trước khi phân tích."); return
        analysis_func(self.df, self.subject_combo.get(), self.plot_frame)

    def handle_distribution(self): self.run_analysis(plot_score_distribution)
    def handle_classification(self): self.run_analysis(plot_classification_pie)
    def handle_province_comparison(self): self.run_analysis(plot_province_comparison)

if __name__ == "__main__":
    app = App()
    app.mainloop()