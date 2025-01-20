#TABLO 1
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sqlite3
import pandas as pd
import os

class DBComparisonApp:
    def __init__(self, master, db1_path, db2_path, ders_kodu, ders_adi):
        self.master = master
        self.master.title("Program-Ders Çıktıları İlişki Matrisi")
        self.master.geometry("1200x800")

        # Renkler
        self.bg_color = "#f0f0f0"
        self.primary_color = "#dd91b9"
        self.secondary_color = "#93d2e3"
        self.accent_color = "#d7b3f0"

        self.master.configure(bg=self.bg_color)

        self.style = ttk.Style(master)
        self.style.theme_use('clam')

        self.style.configure('TFrame', background=self.bg_color)
        self.style.configure('TLabel', background=self.bg_color, foreground='black')
        self.style.configure('TButton', background=self.secondary_color, foreground='black', font=('Arial', 10))
        self.style.map('TButton', background=[('active', self.accent_color)])
        self.style.configure('TEntry', fieldbackground='white', foreground='black')

        self.db1_path = db1_path
        self.db2_path = db2_path
        self.ders_kodu = ders_kodu  #ders kodu kaydet
        self.ders_adi = ders_adi    #ders adı kaydet

        self.main_frame = ttk.Frame(master, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.matrix_frame = ttk.Frame(self.main_frame)
        self.matrix_frame.pack(fill=tk.BOTH, expand=True)

        #ilişkiler için giriş alanlarını saklayan sözlük
        self.relation_entries = {}
        self.relation_labels = {}

        #excelden içeri aktarma ve excele aktar butonu
        ttk.Button(self.main_frame, text="Excel'den Aktar", command=self.import_from_excel).pack(pady=10)
        ttk.Button(self.main_frame, text="Excel'e Aktar", command=self.export_to_excel).pack(pady=10)

        self.load_data()

        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.matrix_frame.grid_columnconfigure(tuple(range(12)), weight=1)
        self.matrix_frame.grid_rowconfigure(tuple(range(1, 12)), weight=1)

    #veritabanından tabloları okuma
    def get_tables(self, db_path, is_first_db=True):
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            # İlk veritabanı ders verileri, diğeri program verileri
            table_name = "ders_verileri" if is_first_db else "program_verileri"
            cursor.execute(f"SELECT sira_no, aciklama FROM {table_name} ORDER BY sira_no")
            tables = cursor.fetchall()
            conn.close()
            return tables

        except Exception as e:
            messagebox.showerror("Hata", f"Veritabanı okuma hatası: {str(e)}")
            return []

    #verileri yükleme ve matris oluşturma
    def load_data(self):
        if os.path.exists(self.db1_path) and os.path.exists(self.db2_path):
            self.update_matrix()
        else:
            messagebox.showerror("Hata", "Veritabanı dosyaları bulunamadı!")

    #belirli bir satırın ilişki değerini hesaplama
    def calculate_relation(self, row_no):
        total = 0
        count = 0
        #girilen değerleri kontrol ederek toplam ve sayıyı bulma
        for (r, c), entry in self.relation_entries.items():
            if r == row_no and entry.get() and entry.get() != '-':
                try:
                    value = float(entry.get().replace(',', '.'))  #hem nokta hem virgül desteklemek için
                    if 1 <= c <= 10:  #sadece 1-10 arası program çıktıları için hesaplama yap
                        total += value
                        count += 1
                except ValueError:
                    pass
        return total / count if count > 0 else 0

    #giriş alanında değişiklik olduğunda tetiklenen fonksiyon
    def on_entry_change(self, event):
        entry = event.widget
        try:
            value = float(entry.get().replace(',', '.'))  #hem nokta hem virgül desteklemek için
            if not (0 <= value <= 1):
                raise ValueError
        except ValueError:
            entry.delete(0, tk.END)
            entry.insert(0, '-')
        self.update_relation_labels()

    #tüm ilişki değeri etiketlerini güncelle
    def update_relation_labels(self):
        for row_no, label in self.relation_labels.items():
            relation = self.calculate_relation(row_no)
            label.config(text=f"{relation:.2f}")

    #matrisi yeniden oluştur
    def update_matrix(self):
        for widget in self.matrix_frame.winfo_children():
            widget.destroy()

        #program ve ders çıktılarını getir
        rows = self.get_tables(self.db2_path, False)  #program çıktıları
        cols = self.get_tables(self.db1_path, True)   #ders çıktıları

        #matris başlığı
        ttk.Label(self.matrix_frame, text="Program Çıktısı/\nDers Çıktısı",
                font=('Arial', 10, 'bold')).grid(row=0, column=0, padx=5, pady=5)

        #ders çıktıları (sütun)
        for col_idx, (no, desc) in enumerate(cols, 1):
            ttk.Label(self.matrix_frame, text=str(no),
                    font=('Arial', 10, 'bold')).grid(row=0, column=col_idx, padx=5, pady=5)

        #ilişki deperi başlığı
        ttk.Label(self.matrix_frame, text="İlişki\nDeğeri",
                font=('Arial', 10, 'bold')).grid(row=0, column=len(cols)+1, padx=5, pady=5)

        self.relation_labels = {}

        #program çıktıları (satır)
        for row_idx, (row_no, row_desc) in enumerate(rows, 1):
            ttk.Label(self.matrix_frame, text=str(row_no),
                    font=('Arial', 10, 'bold')).grid(row=row_idx, column=0, padx=5, pady=5)

            for col_idx, (col_no, col_desc) in enumerate(cols, 1):
                #girşi alanlarını oluşturma
                entry = ttk.Entry(self.matrix_frame, width=5)
                entry.insert(0, '-')
                entry.grid(row=row_idx, column=col_idx, padx=2, pady=2, sticky="nsew")
                entry.bind('<KeyRelease>', self.on_entry_change)
                entry.bind('<FocusIn>', lambda e: e.widget.delete(0, tk.END))
                self.relation_entries[(row_no, col_no)] = entry

            #ilişki değperi etiketi
            relation_label = ttk.Label(self.matrix_frame, text="0.00")
            relation_label.grid(row=row_idx, column=len(cols)+1, padx=5, pady=5, sticky="nsew")
            self.relation_labels[row_no] = relation_label

    #matris verilerini excele aktarma
    def export_to_excel(self):
        try:
            matrix_data = []
            rows = self.get_tables(self.db2_path, False)
            cols = self.get_tables(self.db1_path, True)

            for row_no, row_desc in rows:
                row_data = {'Program Çıktısı': row_no, 'Açıklama': row_desc}
                for col_no, col_desc in cols:
                    entry = self.relation_entries.get((row_no, col_no))
                    value = entry.get() if entry and entry.get() != '-' else '0'
                    row_data[f'{col_no}'] = value
                row_data['İlişki Değeri'] = f"{self.calculate_relation(row_no):.2f}"
                matrix_data.append(row_data)

            #excel dosya kaydetme
            df = pd.DataFrame(matrix_data)
            current_dir = os.path.dirname(os.path.abspath(__file__))
            #yeni dosya adı formatı
            filename = f"{self.ders_kodu}_{self.ders_adi}_tablo_1.xlsx"
            save_path = os.path.join(current_dir, filename)
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Başarılı", f"İlişki matrisi {filename} olarak kaydedildi!")

        except Exception as e:
            messagebox.showerror("Hata", f"Excel kaydetme hatası: {str(e)}")

    #excelden veri içeri aktarımı için
    def import_from_excel(self):
        try:
            file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
            if not file_path:
                return

            df = pd.read_excel(file_path)
            rows = self.get_tables(self.db2_path, False)
            cols = self.get_tables(self.db1_path, True)

            for row_no, row_desc in rows:
                for col_no, col_desc in cols:
                    #excel sütun adlarının tam olarak eşleştiğinden emin olmak için
                    if str(col_no) in df.columns and 'Program Çıktısı' in df.columns:
                        value = df.loc[df['Program Çıktısı'] == row_no, str(col_no)].values
                        if len(value) > 0:
                            entry = self.relation_entries.get((row_no, col_no))
                            if entry:
                                entry.delete(0, tk.END)
                                entry.insert(0, str(value[0]))

            self.update_relation_labels()  # Otomatik hesaplama
            messagebox.showinfo("Başarılı", "Excel verileri başarıyla yüklendi ve ilişki değerleri hesaplandı!")

        except Exception as e:
            messagebox.showerror("Hata", f"Excel'den veri okuma hatası: {str(e)}")

def run_tablo_1_app(master, db1_path, db2_path, ders_kodu, ders_adi):
    # Ana pencere yerine yeni bir pencere oluştur
    tablo1_window = tk.Toplevel(master)
    tablo1_window.title("Program-Ders Çıktıları İlişki Matrisi")
    tablo1_window.geometry("1200x800")

    # Renkler
    bg_color = "#f0f0f0"
    tablo1_window.configure(bg=bg_color)

    app = DBComparisonApp(tablo1_window, db1_path, db2_path, ders_kodu, ders_adi)

if __name__ == "__main__":
    # Bu blok artık doğrudan çalıştırılmayacak, sadece ana uygulamadan çağrılacak.
    pass