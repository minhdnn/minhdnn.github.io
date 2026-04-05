import os
import re
import pandas as pd
import xml.etree.ElementTree as ET
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk

# Thiết lập giao diện
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class ProInvoiceApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Trích xuất Hóa đơn XML chi tiết (by Minh Do)")
        self.geometry("1100x800")

        # Khai báo các cột dữ liệu (Thêm "Đường dẫn tới folder")
        self.header_fields = [
            "Số hóa đơn", "Ký hiệu", "Ngày lập",
            "Mã số thuế bán", "Tên đơn vị bán",
            "Tên đơn vị mua", "Tổng tiền thanh toán hóa đơn",
            "Đường dẫn file XML", "Đường dẫn tới folder"
        ]

        # Khai báo các cột chi tiết
        self.item_fields = [
            "STT", "Tên hàng hóa", "Đơn vị tính",
            "Số lượng", "Đơn giá", "Thành tiền chưa thuế",
            "Thuế suất", "Tiền thuế mặt hàng", "Ghi chú thuế"
        ]

        self.selected_fields = {}
        self.current_data = []
        self.selected_files_list = []

        self.init_ui()

    def init_ui(self):
        # --- Vùng 1: Chọn file/thư mục ---
        frame_top = ctk.CTkFrame(self, fg_color="transparent")
        frame_top.pack(pady=5, fill="x", padx=20)

        ctk.CTkLabel(
            frame_top,
            text="Đường dẫn file/thư mục XML đã chọn:",
            font=("Arial", 12, "bold")
        ).pack(anchor="w")

        path_frame = ctk.CTkFrame(frame_top, fg_color="transparent")
        path_frame.pack(fill="x", pady=5)

        self.entry_path = ctk.CTkEntry(path_frame, width=500, state="disabled")
        self.entry_path.pack(side="left", padx=(0, 10), fill="x", expand=True)

        ctk.CTkButton(
            path_frame,
            text="📁 Chọn 1 File",
            width=120,
            command=self.select_file
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            path_frame,
            text="📂 Chọn Thư mục",
            width=120,
            command=self.select_folder
        ).pack(side="left")

        # --- Vùng 2: Chọn trường dữ liệu ---
        frame_mid = ctk.CTkFrame(self, fg_color="transparent")
        frame_mid.pack(pady=5, fill="x", padx=20)

        ctk.CTkLabel(
            frame_mid,
            text="Chọn cột dữ liệu xuất ra Excel:",
            font=("Arial", 12, "bold")
        ).pack(anchor="w", pady=2)

        columns_frame = ctk.CTkFrame(frame_mid, fg_color="transparent")
        columns_frame.pack(fill="x")

        frame_left = ctk.CTkScrollableFrame(
            columns_frame,
            label_text="Thông tin chung",
            height=120
        )
        frame_left.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        # Mặc định uncheck cho 2 cột đường dẫn
        for name in self.header_fields:
            is_checked = False if name in ["Đường dẫn file XML", "Đường dẫn tới folder"] else True
            var = ctk.BooleanVar(value=is_checked)
            ctk.CTkCheckBox(frame_left, text=name, variable=var).pack(anchor="w", padx=10, pady=2)
            self.selected_fields[name] = var

        frame_right = ctk.CTkScrollableFrame(
            columns_frame,
            label_text="Chi tiết hàng hóa",
            height=120
        )
        frame_right.pack(side="right", fill="both", expand=True, padx=(5, 0))
        
        for name in self.item_fields:
            var = ctk.BooleanVar(value=True)
            ctk.CTkCheckBox(frame_right, text=name, variable=var).pack(anchor="w", padx=10, pady=2)
            self.selected_fields[name] = var

        # --- Vùng 3: Nút điều khiển ---
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=10)

        self.btn_preview = ctk.CTkButton(
            btn_frame,
            text="👀 XEM TRƯỚC",
            font=("Arial", 14, "bold"),
            fg_color="#17a2b8",
            hover_color="#138496",
            height=40,
            command=self.preview_data
        )
        self.btn_preview.pack(side="left", padx=10)

        self.btn_export = ctk.CTkButton(
            btn_frame,
            text="🚀 XUẤT EXCEL",
            font=("Arial", 14, "bold"),
            fg_color="#28a745",
            hover_color="#218838",
            height=40,
            command=self.export_excel
        )
        self.btn_export.pack(side="left", padx=10)

        # --- Vùng 4: Progress Bar & Trạng thái ---
        self.progress_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.progress_frame.pack(fill="x", padx=20, pady=(0, 10))

        self.status_label = ctk.CTkLabel(
            self.progress_frame,
            text="Sẵn sàng nạp dữ liệu.",
            font=("Arial", 12, "italic"),
            text_color="gray"
        )
        self.status_label.pack(anchor="w")

        self.progress_bar = ctk.CTkProgressBar(
            self.progress_frame,
            mode="determinate",
            height=10
        )
        self.progress_bar.pack(fill="x", pady=(2, 0))
        self.progress_bar.set(0)

        # --- Vùng 5: Bảng Preview (Treeview) ---
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        tree_scroll_y = ttk.Scrollbar(self.tree_frame)
        tree_scroll_y.pack(side="right", fill="y")
        tree_scroll_x = ttk.Scrollbar(self.tree_frame, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")

        self.tree = ttk.Treeview(
            self.tree_frame,
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set
        )
        self.tree.pack(fill="both", expand=True)
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)

    def update_entry_path(self, text):
        self.entry_path.configure(state="normal")
        self.entry_path.delete(0, 'end')
        self.entry_path.insert(0, text)
        self.entry_path.configure(state="disabled")

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
        if path:
            self.selected_files_list = [path]
            self.update_entry_path(path)
            self.progress_bar.set(0)
            self.status_label.configure(
                text="Đã nạp 1 file. Bấm Xem trước hoặc Xuất Excel để bắt đầu.",
                text_color="green"
            )

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if not folder_path:
            return

        found_files = []
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.xml'):
                    found_files.append(os.path.join(root, file))

        if not found_files:
            messagebox.showwarning("Thông báo", "Không tìm thấy file XML nào trong thư mục này!")
            return

        self.show_confirmation_dialog(folder_path, found_files)

    def show_confirmation_dialog(self, folder_path, files):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Xác nhận danh sách file")
        dialog.geometry("700x450")
        dialog.transient(self)
        dialog.grab_set()

        lbl = ctk.CTkLabel(
            dialog,
            text=f"🔍 Tìm thấy {len(files)} file XML. Vui lòng kiểm tra:",
            font=("Arial", 14, "bold")
        )
        lbl.pack(pady=(15, 5))

        textbox = ctk.CTkTextbox(dialog, width=650, height=300)
        textbox.pack(pady=10, padx=20)

        files_text = "\n".join(files)
        textbox.insert("0.0", files_text)
        textbox.configure(state="disabled")

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=10)

        def confirm_action():
            self.selected_files_list = files
            self.update_entry_path(folder_path)
            self.progress_bar.set(0)
            self.status_label.configure(
                text=f"Đã nạp danh sách {len(files)} file. Bấm Xem trước hoặc Xuất Excel để bắt đầu.",
                text_color="green"
            )
            dialog.destroy()

        def cancel_action():
            dialog.destroy()

        ctk.CTkButton(
            btn_frame,
            text="✅ Xác nhận & Nạp",
            fg_color="#28a745",
            hover_color="#218838",
            command=confirm_action
        ).pack(side="left", padx=15)

        ctk.CTkButton(
            btn_frame,
            text="❌ Hủy",
            fg_color="#dc3545",
            hover_color="#c82333",
            command=cancel_action
        ).pack(side="left", padx=15)

    def parse_xml_pro(self, xml_path):
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            for el in root.iter():
                if '}' in el.tag:
                    el.tag = el.tag.split('}', 1)[1]

            tt_chung = root.find(".//DLHDon/TTChung")
            if tt_chung is None:
                tt_chung = root.find(".//TTChung")

            n_ban = root.find(".//NBan")
            n_mua = root.find(".//NMua")
            t_toan = root.find(".//TToan")

            normalized_path = os.path.normpath(xml_path)

            base_data = {
                "Số hóa đơn": tt_chung.findtext("SHDon") if tt_chung is not None else "",
                "Ký hiệu": tt_chung.findtext("KHHDon") if tt_chung is not None else "",
                "Ngày lập": tt_chung.findtext("NLap") if tt_chung is not None else "",
                "Mã số thuế bán": n_ban.findtext("MST") if n_ban is not None else "",
                "Tên đơn vị bán": n_ban.findtext("Ten") if n_ban is not None else "",
                "Tên đơn vị mua": n_mua.findtext("Ten") if n_mua is not None else "",
                "Tổng tiền thanh toán hóa đơn": t_toan.findtext("TgTTTBSo") if t_toan is not None else "0",
                "Đường dẫn file XML": normalized_path,
                "Đường dẫn tới folder": os.path.dirname(normalized_path)  # Lấy thư mục chứa file
            }

            rows = []
            
            # Kiểm tra xem có cột "Chi tiết hàng hóa" nào đang được chọn không
            has_item_fields_selected = any(self.selected_fields[field].get() for field in self.item_fields)
            
            dshhdvu = root.find(".//DSHHDVu")
            
            # Nếu người dùng CÓ chọn xuất chi tiết hàng hóa VÀ hóa đơn CÓ chi tiết hàng hóa
            if has_item_fields_selected and dshhdvu is not None and len(dshhdvu.findall("HHDVu")) > 0:
                for hhdvu in dshhdvu.findall("HHDVu"):
                    row_data = base_data.copy()

                    row_data["STT"] = hhdvu.findtext("STT")
                    row_data["Tên hàng hóa"] = hhdvu.findtext("THHDVu")
                    row_data["Đơn vị tính"] = hhdvu.findtext("DVTinh")
                    row_data["Số lượng"] = hhdvu.findtext("SLuong")
                    row_data["Đơn giá"] = hhdvu.findtext("DGia")

                    th_tien = hhdvu.findtext("ThTien") or "0"
                    row_data["Thành tiền chưa thuế"] = th_tien

                    tsuat = hhdvu.findtext("TSuat")
                    row_data["Thuế suất"] = tsuat

                    # BIẾN GHI CHÚ
                    ghi_chu_thue = ""

                    if tsuat in ["KCT", "KKKNT", "0%"]:
                        row_data["Tiền thuế mặt hàng"] = tsuat
                        ghi_chu_thue = "Hàng KCT/KKKNT"
                    else:
                        tien_thue = "0"
                        found_tax = False

                        # 1. Tìm thẻ <TThue> công khai
                        tthue_node = hhdvu.find("TThue")
                        if tthue_node is not None and tthue_node.text:
                            tien_thue = tthue_node.text
                            found_tax = True
                            ghi_chu_thue = "Lấy từ XML (Thẻ chuẩn)"
                        else:
                            # 2. Tìm thẻ bị ẩn trong <TTKhac> -> <TTin>
                            for ttin in hhdvu.findall(".//TTin"):
                                truong = ttin.findtext("TTruong")
                                if truong in ["Tiền thuế", "VATAmount", "TienThue"]:
                                    dlieu = ttin.findtext("DLieu")
                                    if dlieu:
                                        tien_thue = dlieu
                                        found_tax = True
                                        ghi_chu_thue = "Lấy từ XML (Thẻ ẩn)"
                                        break

                        # 3. App tự động tính toán
                        if not found_tax and tsuat:
                            try:
                                match = re.search(r'(\d+(\.\d+)?)', tsuat)
                                if match and th_tien:
                                    rate = float(match.group(1))
                                    tien_thue = str(round(float(th_tien) * (rate / 100)))
                                    ghi_chu_thue = "App tự động tính"
                                else:
                                    ghi_chu_thue = "Không thể tính thuế"
                            except Exception:
                                ghi_chu_thue = "Lỗi khi tự tính"

                        row_data["Tiền thuế mặt hàng"] = tien_thue

                    row_data["Ghi chú thuế"] = ghi_chu_thue

                    filtered_row = {
                        k: v for k, v in row_data.items()
                        if k in self.selected_fields and self.selected_fields[k].get()
                    }
                    rows.append(filtered_row)
            
            # Nếu người dùng KHÔNG chọn cột chi tiết nào HOẶC hóa đơn trống rỗng mục hàng hóa
            else:
                filtered_row = {
                    k: v for k, v in base_data.items()
                    if k in self.selected_fields and self.selected_fields[k].get()
                }
                rows.append(filtered_row)

            return rows
        except Exception as e:
            print(f"Lỗi đọc {xml_path}: {e}")
            return []

    def load_data(self):
        if not self.selected_files_list:
            messagebox.showwarning("Cảnh báo", "Bạn chưa chọn file hoặc thư mục nào!")
            return False

        total_files = len(self.selected_files_list)
        self.current_data = []
        self.progress_bar.set(0)

        for index, f in enumerate(self.selected_files_list):
            self.status_label.configure(
                text=f"Đang đọc dữ liệu... {index + 1}/{total_files} file",
                text_color="#17a2b8"
            )
            self.current_data.extend(self.parse_xml_pro(f))

            progress = (index + 1) / total_files
            self.progress_bar.set(progress)
            self.update_idletasks()

        if not self.current_data:
            self.status_label.configure(
                text="Thất bại: Không có dữ liệu hợp lệ.",
                text_color="red"
            )
            messagebox.showerror("Lỗi", "Không có dữ liệu hợp lệ nào được quét ra!")
            return False

        self.status_label.configure(
            text=f"Thành công: Đã quét xong {total_files} file.",
            text_color="green"
        )
        return True

    def preview_data(self):
        if not self.load_data():
            return

        self.tree.delete(*self.tree.get_children())

        if self.current_data:
            self.status_label.configure(
                text="Đang dựng bảng xem trước...",
                text_color="#17a2b8"
            )
            self.update_idletasks()

            columns = list(self.current_data[0].keys())
            self.tree["columns"] = columns
            self.tree["show"] = "headings"

            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=120, anchor="w")

            for row in self.current_data:
                self.tree.insert("", "end", values=[row.get(col, "") for col in columns])

            self.status_label.configure(
                text="Đã hiển thị bảng xem trước thành công.",
                text_color="green"
            )

    def export_excel(self):
        if not self.load_data():
            return

        self.status_label.configure(
            text="Đang định dạng và lưu ra file Excel...",
            text_color="#17a2b8"
        )
        self.update_idletasks()

        df = pd.DataFrame(self.current_data)

        # Chuyển kiểu số cho các cột numeric
        numeric_cols = [
            "Số lượng", "Đơn giá",
            "Thành tiền chưa thuế",
            "Tiền thuế mặt hàng",
            "Tổng tiền thanh toán hóa đơn"
        ]
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if save_path:
            try:
                # Fix cho macOS: chỉ định engine openpyxl
                df.to_excel(save_path, index=False, engine="openpyxl")
                self.status_label.configure(
                    text=f"Hoàn thành: Đã xuất file Excel tại {save_path}",
                    text_color="green"
                )
                messagebox.showinfo(
                    "Thành công",
                    f"Đã xuất {len(self.current_data)} dòng ra:\n{save_path}"
                )
            except Exception as e:
                self.status_label.configure(
                    text="Lỗi: Không thể lưu file Excel.",
                    text_color="red"
                )
                messagebox.showerror(
                    "Lỗi khi xuất",
                    "Có lỗi xảy ra khi lưu file.\n"
                    "- Kiểm tra đã cài thư viện openpyxl chưa (pip install openpyxl)\n"
                    "- Thử lưu ra Desktop/Documents thay vì thư mục hệ thống đặc biệt trên macOS.\n\n"
                    f"Chi tiết kỹ thuật:\n{e}"
                )
        else:
            self.status_label.configure(
                text="Đã hủy lưu file Excel.",
                text_color="orange"
            )

if __name__ == "__main__":
    app = ProInvoiceApp()
    app.mainloop()