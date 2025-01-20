#TABLO 4
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import sqlite3
import os

class StudentOutputCalculator:
    def __init__(self, root, ders_kodu, ders_adi):
        #uygulama penceresinin başlat
        self.root = root
        self.root.title("Öğrenci Ders Çıktıları Hesaplama")
        self.root.geometry("1200x800")

        # Renkler
        self.bg_color = "#f0f0f0"
        self.primary_color = "#dd91b9"
        self.secondary_color = "#93d2e3"
        self.accent_color = "#d7b3f0"

        self.root.configure(bg=self.bg_color)

        self.style = ttk.Style(root)
        self.style.theme_use('clam')

        self.style.configure('TNotebook.Tab', background=self.secondary_color, foreground='black', font=('Arial', 10))
        self.style.map('TNotebook.Tab', background=[("selected", self.accent_color)], foreground=[("selected", 'black')])
        self.style.configure('TFrame', background=self.bg_color)
        self.style.configure('TLabel', background=self.bg_color, foreground='black')
        self.style.configure('TButton', background=self.secondary_color, foreground='black', font=('Arial', 10))
        self.style.map('TButton', background=[('active', self.accent_color)])
        self.style.configure('Treeview.Heading', font=('Arial', 10, 'bold'))
        self.style.configure('Treeview', background="white", foreground="black", fieldbackground="white")

        self.ders_kodu = ders_kodu
        self.ders_adi = ders_adi

        #tablar için notebook widgetı oluşturma
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        #tabloları depolayacak liste
        self.trees = []

        #veri yükleme işlemi için buton
        ttk.Button(root, text="Verileri Yükle ve Hesapla", command=self.load_and_calculate).pack(pady=5)

        #tüm tabloları excele kaydetmek için buton
        ttk.Button(root, text="Tüm Tabloları Excel'e Kaydet", command=self.save_all_to_excel).pack(pady=5)

        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

    def load_and_calculate(self):
        self.load_data()
        if hasattr(self, 'student_grades') and hasattr(self, 'output_matrix') and hasattr(self, 'ders_ciktilari'):
            self.calculate_and_create_tabs()

    def load_data(self):
        try:
            #öğrenci notlarını excel dosyasından okuma
            not_file_name = f"{self.ders_kodu}_not.xlsx"
            self.student_grades = pd.read_excel(not_file_name)

            #öğrenci numaralarını tamsayıya dönüştürme
            if 'Ogrenci_No' in self.student_grades.columns:
                self.student_grades['Ogrenci_No'] = self.student_grades['Ogrenci_No'].astype(str)
            else:
                messagebox.showerror("Hata", f"Öğrenci notları dosyasında 'Ogrenci_No' sütunu bulunamadı.")
                return

            #ders çıktıları matrisini excel dosyasından okuma
            matrix_file_name = f"{self.ders_kodu}_{self.ders_adi.replace(' ', '_')}_tablo_2_3.xlsx"
            self.output_matrix = pd.read_excel(matrix_file_name, sheet_name='Tablo 3')

            #ders çıktıları veritabanından alma
            self.ders_ciktilari = self.get_ders_ciktilari_from_db()

        except FileNotFoundError as e:
            messagebox.showerror("Hata", f"Dosya bulunamadı: {e.filename}. Lütfen dosyaların doğru konumda olduğunu kontrol edin.")
        except Exception as e:
            messagebox.showerror("Hata", f"Dosya okuma hatası: {e}")

    def get_ders_ciktilari_from_db(self):
        try:
            #veritabanına bağlanma ve ders çıktılarının alınması
            conn = sqlite3.connect('ders_ciktilari.db')
            cursor = conn.cursor()

            #sql sorgusu ile ders çıktıları alma
            cursor.execute("SELECT aciklama FROM ders_verileri ORDER BY sira_no")
            results = [row[0] for row in cursor.fetchall()]

            conn.close()
            return results
        except Exception as e:
            messagebox.showerror("Hata", f"Ders çıktıları veritabanından alınamadı: {e}")
            return []

    def calculate_and_create_tabs(self):
        try:
            #mevcut sekmeleri temizle
            for tab in self.notebook.tabs():
                self.notebook.forget(tab)
            self.trees = []  #önceki tabloları temizle

            #her öğrenci için hesaplama yapıp tablo oluşturma !!
            for idx, student in self.student_grades.iterrows():
                student_no = student['Ogrenci_No']
                tab_frame = ttk.Frame(self.notebook)
                self.notebook.add(tab_frame, text=f"Öğrenci {student_no}")

                #tablo başlıkları
                columns = ("Ders Çıktısı", "Ödev1", "Ödev2", "Quiz", "Vize", "Final", "TOPLAM", "MAX", "% Başarı")
                tree = ttk.Treeview(tab_frame, columns=columns, show="headings")

                #kolon başlıklarını ayarlama
                for col in columns:
                    tree.heading(col, text=col)
                    tree.column(col, width=100, anchor="center")

                tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

                #hesaplama ve tablonun oluşturma
                self.calculate_and_display(tree, student)

                #tabloyu listeye ekleyerek daha sonra kaydetmek iiçin sakla
                self.trees.append((student_no, tree))

                tab_frame.grid_columnconfigure(0, weight=1)
                tab_frame.grid_rowconfigure(0, weight=1)

        except Exception as e:
            messagebox.showerror("Hata", f"Hesaplama hatası: {e}")

    def calculate_and_display(self, tree, student):
        try:
            #önceden mevcut tüm verileri tablodan temizle
            for item in tree.get_children():
                tree.delete(item)

            column_mapping = {
                'Ödev1': 'Ödev1',
                'Ödev2': 'Ödev2',
                'Quiz': 'Quiz',
                'Vize': 'Vize',
                'Final': 'Final'
            }

            #tablo 3 oku
            matrix_file_name = f"{self.ders_kodu}_{self.ders_adi.replace(' ', '_')}_tablo_2_3.xlsx"
            table3_df = pd.read_excel(matrix_file_name, sheet_name='Tablo 3')

            #her ders çıktısı için hesaplama yapma
            for idx, output_row in self.output_matrix.iterrows():
                output_no = output_row['Ders Çıktısı']
                ders_cikti = self.ders_ciktilari[idx] if idx < len(self.ders_ciktilari) else "Bilinmiyor"

                #öğrenci notları ile hesaplama yapma
                total = 0
                for matrix_col, grade_col in column_mapping.items():
                    if matrix_col in self.output_matrix.columns:
                        weight = float(output_row[matrix_col])
                        grade = float(student[grade_col])
                        total += weight * grade

                #max değeri alıp başarı oranını hesaplama
                max_value = float(table3_df[table3_df['Ders Çıktısı'] == output_no]['TOPLAM'].iloc[0]) * 100
                success_rate = (total / max_value) * 100 if max_value > 0 else 0

                #hesaplanan verileri tabloya ekleme
                tree.insert('', 'end', values=(
                    ders_cikti, student['Ödev1'], student['Ödev2'], student['Quiz'],
                    student['Vize'], student['Final'], f"{total:.1f}",
                    f"{max_value:.1f}", f"{success_rate:.1f}"
                ))

        except Exception as e:
            messagebox.showerror("Hata", f"Hesaplama veya görüntüleme hatası: {e}")

    def save_all_to_excel(self):
        try:
            #mevcut çalışma dizinini al
            current_dir = os.getcwd()
            file_name = f"{self.ders_kodu}_tablo_4.xlsx"
            file_path = os.path.join(current_dir, file_name)

            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                for student_no, tree in self.trees:
                    data = []
                    for item in tree.get_children():
                        data.append(tree.item(item)['values'])

                    #dataframe oluşturma
                    df = pd.DataFrame(data, columns=[
                        "Ders Çıktısı", "Ödev1", "Ödev2", "Quiz", "Vize", "Final",
                        "TOPLAM", "MAX", "% Başarı"
                    ])

                    #excele kaydet
                    df.to_excel(writer, sheet_name=f"{student_no}", index=False)

            messagebox.showinfo("Başarılı", f"Tüm öğrenci tabloları {file_name} dosyasına kaydedildi!")

        except Exception as e:
            messagebox.showerror("Hata", f"Excel'e kaydetme hatası: {e}")

def run_tablo_4_gui(main_window, ders_kodu, ders_adi):
    tablo4_window = tk.Toplevel(main_window)
    tablo4_window.title("Öğrenci Ders Çıktıları Hesaplama")
    tablo4_window.geometry("1200x800")

    # Renkler
    bg_color = "#f0f0f0"
    tablo4_window.configure(bg=bg_color)

    app = StudentOutputCalculator(tablo4_window, ders_kodu, ders_adi)

if __name__ == "__main__":
    pass