#TABLO 5
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os

class StudentOutputCalculator:
    def __init__(self, root, ders_kodu, ders_adi):
        self.root = root
        self.root.title(f"{ders_kodu} - Program Çıktısı Başarı Oranları (Tablo 5)")
        self.root.geometry("800x600")

        # Renkler
        self.bg_color = "#f0f0f0"
        self.primary_color = "#dd91b9"
        self.secondary_color = "#93d2e3"
        self.accent_color = "#d7b3f0"

        self.root.configure(bg=self.bg_color)

        self.style = ttk.Style(root)
        self.style.theme_use('clam')

        self.style.configure('TFrame', background=self.bg_color)
        self.style.configure('TLabel', background=self.bg_color, foreground='black')
        self.style.configure('TButton', background=self.secondary_color, foreground='black', font=('Arial', 10))
        self.style.map('TButton', background=[('active', self.accent_color)])

        self.ders_kodu = ders_kodu
        self.ders_adi = ders_adi

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Butonu burada, process_all_students tanımlandıktan sonra oluşturuyoruz.
        self.process_button = ttk.Button(
            root, text="Tüm Tablo 5'leri Oluştur", command=self.process_all_students
        )
        self.process_button.pack(pady=20)

        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

    def safe_float_convert(self, value):
        try:
            if isinstance(value, str):
                value = value.replace(',', '.') #hem nokta hem virgül için
            return float(value)
        except (ValueError, TypeError):
            return 0.0

    def get_student_numbers(self):
        try:
            df = pd.read_excel(f"{self.ders_kodu}_not.xlsx", dtype=str) # Öğrenci numaralarını string olarak oku
            return df.iloc[:, 0].tolist()
        except Exception as e:
            messagebox.showerror("Hata", f"Öğrenci notları dosyası okuma hatası: {str(e)}")
            return []

    def read_tablo1(self):
        try:
            df = pd.read_excel(f"{self.ders_kodu}_{self.ders_adi.replace(' ', '_')}_tablo_1.xlsx")
            # Sayısal sütunları ve ilişki değerini garantiye almak için
            numeric_data_df = df.iloc[:, 2:-1].apply(pd.to_numeric, errors='coerce').fillna(0)
            iliski_degerleri = pd.to_numeric(df.iloc[:, -1], errors='coerce').fillna(0).tolist()
            return numeric_data_df.values.tolist(), iliski_degerleri
        except Exception as e:
            messagebox.showerror("Hata", f"Tablo 1 okuma hatası: {str(e)}")
            return [], []

    def process_all_students(self):
        tablo1_values, iliski_degerleri = self.read_tablo1()
        if not tablo1_values:
            return

        student_numbers = self.get_student_numbers()
        if not student_numbers:
            return

        # Mevcut sekmeleri temizle
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)

        for student_no in student_numbers:
            try:
                tablo5_data = self.create_tablo_5_data(student_no, tablo1_values, iliski_degerleri)
                if tablo5_data is not None and not tablo5_data.empty:
                    self.display_tablo_5(student_no, tablo5_data, tablo5_data.columns.tolist()) # Sütun başlıklarını gönder
                else:
                    messagebox.showwarning("Uyarı", f"{student_no} için Tablo 5 verisi oluşturulamadı veya boş.")
            except Exception as e:
                messagebox.showerror("Hata", f"İşlem hatası - Öğrenci: {student_no} - Hata: {str(e)}")

    def create_tablo_5_data(self, student_no, tablo1_values, iliski_degerleri):
        try:
            tablo4_filename = f"{self.ders_kodu}_tablo_4.xlsx"
            student_sheet_name = str(student_no)
            tablo_4_df = pd.read_excel(tablo4_filename, sheet_name=student_sheet_name)

            if "% Başarı" not in tablo_4_df.columns:
                raise KeyError(f"{student_sheet_name} sayfasında '% Başarı' sütunu bulunamadı.")

            basari_yuzdeleri = pd.to_numeric(tablo_4_df["% Başarı"], errors='coerce').fillna(0).tolist()
            tablo_5_data = []

            for row_idx, (tablo1_row, iliski_degeri) in enumerate(zip(tablo1_values, iliski_degerleri), start=1):
                row_data = {"Prg Çıktı": row_idx}
                satir_degerleri = []
                for ders_idx, basari in enumerate(basari_yuzdeleri):
                    carpim = basari * tablo1_row[ders_idx]
                    row_data[f"Ders çıktısı {ders_idx+1}"] = carpim
                    satir_degerleri.append(carpim)

                satir_toplami = sum(satir_degerleri)
                basari_orani = (satir_toplami / 5) / iliski_degeri if iliski_degeri != 0 else 0
                row_data["Başarı Oranı"] = basari_orani
                tablo_5_data.append(row_data)

            tablo_5_df = pd.DataFrame(tablo_5_data)
            column_mapping = {"Prg Çıktı": "Prg Çıktı"} # Sabit sütun başlığı
            for i, yuzde in enumerate(basari_yuzdeleri):
                column_mapping[f"Ders çıktısı {i+1}"] = f"{yuzde:.1f}%"
            column_mapping["Başarı Oranı"] = "Başarı Oranı"
            tablo_5_df.rename(columns=column_mapping, inplace=True)
            return tablo_5_df
        except FileNotFoundError:
            messagebox.showwarning("Uyarı", f"Tablo 4 dosyası bulunamadı: {student_no}")
            return pd.DataFrame()  # Boş DataFrame döndür
        except Exception as e:
            messagebox.showerror("Hata", f"Tablo 5 veri oluşturma hatası: {str(e)}")
            return pd.DataFrame()

    def display_tablo_5(self, student_no, df, columns):
        tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(tab_frame, text=f"Öğrenci {student_no}")

        tree = ttk.Treeview(tab_frame, show="headings")

        # Sütunları benzersiz ID'lerle tanımla
        tree_columns = ["Prg Cikti"] + [f"ders_ciktisi_{i+1}" for i in range(len(columns) - 2)] + ["Basari Orani"]
        tree['columns'] = tree_columns

        # Başlıkları ve genişlikleri ayarla
        for i, col in enumerate(columns):
            tree.heading(tree_columns[i], text=col)
            tree.column(tree_columns[i], width=100, anchor="center")

        for _, row in df.iterrows():
            # Değerleri stringe çevirirken formatlama
            values = [f"{val:.1f}" if isinstance(val, float) else str(val) for val in row]
            tree.insert("", "end", values=values)

        tree.pack(fill=tk.BOTH, expand=True)

def run_tablo_5_app(main_window, ders_kodu, ders_adi):
    tablo5_window = tk.Toplevel(main_window)
    tablo5_window.title(f"{ders_kodu} - Program Çıktısı Başarı Oranları (Tablo 5)")
    tablo5_window.geometry("800x600")

    bg_color = "#f0f0f0"
    tablo5_window.configure(bg=bg_color)

    app = StudentOutputCalculator(tablo5_window, ders_kodu, ders_adi)
    app.process_all_students()