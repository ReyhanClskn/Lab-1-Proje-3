import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import pandas as pd
import os

class CourseOutputMatrixApp:
    def __init__(self, master, db_path, ders_kodu, ders_adi):
        self.master = master
        self.master.title(f"{ders_kodu} - {ders_adi} Ders Çıktısı Değerlendirme Matrisi")
        self.master.geometry("1400x800")

        #renkler
        self.bg_color = "#f0f0f0"
        self.primary_color = "#dd91b9"
        self.secondary_color = "#93d2e3"
        self.accent_color = "#d7b3f0"

        self.master.configure(bg=self.bg_color)

        self.style = ttk.Style(master)
        self.style.theme_use('clam')

        self.style.configure('TFrame', background=self.bg_color)
        self.style.configure('TLabelframe', background=self.bg_color, borderwidth=2, relief="groove")
        self.style.configure('TLabelframe.Label', foreground='black', font=('Arial', 12, 'bold'))
        self.style.configure('TLabel', background=self.bg_color, foreground='black')
        self.style.configure('TButton', background=self.secondary_color, foreground='black', font=('Arial', 10))
        self.style.map('TButton', background=[('active', self.accent_color)])
        self.style.configure('TEntry', fieldbackground='white', foreground='black', highlightthickness=1, highlightcolor=self.primary_color, borderwidth=1, relief="solid")

        self.db_path = db_path
        self.ders_kodu = ders_kodu
        self.ders_adi = ders_adi
        self.weights_file = f"{ders_kodu}_weights.xlsx"
        self.loaded_weights = {}
        self.weight_entries = {}

        self.assignments = ["Ödev1", "Ödev2", "Quiz", "Vize", "Final"] 

        self.main_frame = ttk.Frame(master, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.weights_frame = ttk.Frame(self.main_frame)
        self.weights_frame.pack(fill=tk.X, pady=(0, 10))

        self.matrix_frame = ttk.Frame(self.main_frame)
        self.matrix_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Separator(self.main_frame, orient='horizontal').pack(fill='x', pady=15)

        self.weighted_matrix_frame = ttk.Frame(self.main_frame)
        self.weighted_matrix_frame.pack(fill=tk.BOTH, expand=True)

        self.relation_entries = {}
        self.weighted_labels = {}
        self.sum_labels = {}
        self.weighted_sum_labels = {}

        self.load_weights_from_excel()
        self.create_weights_row()
        self.create_matrix()
        self.create_weighted_matrix()

        self.load_from_excel()

        ttk.Button(self.main_frame, text="Kaydet", command=self.save_data).pack(pady=10)

        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(3, weight=1)
        self.matrix_frame.grid_columnconfigure(tuple(range(7)), weight=1)
        self.matrix_frame.grid_rowconfigure(tuple(range(1, 15)), weight=1)
        self.weighted_matrix_frame.grid_columnconfigure(tuple(range(7)), weight=1)
        self.weighted_matrix_frame.grid_rowconfigure(tuple(range(1, 15)), weight=1)

    def load_weights_from_excel(self):
        try:
            weights_df = pd.read_excel(self.weights_file)
            if 'Assignment' not in weights_df.columns or 'Weight' not in weights_df.columns:
                messagebox.showwarning("Uyarı", f"'{self.weights_file}' dosyasında 'Assignment' veya 'Weight' sütunları bulunamadı. Varsayılan değerler kullanılacak.")
                for assignment in self.assignments:
                    self.loaded_weights[assignment] = 20
                return

            for index, row in weights_df.iterrows():
                assignment_name = row['Assignment']
                weight_value = row['Weight']
                if assignment_name in self.assignments:
                    try:
                        self.loaded_weights[assignment_name] = int(weight_value)
                    except ValueError:
                        messagebox.showwarning("Uyarı", f"'{self.weights_file}' dosyasında '{assignment_name}' için geçersiz ağırlık değeri. Varsayılan değer (20) kullanılacak.")
                        self.loaded_weights[assignment_name] = 20
                else:
                    messagebox.showwarning("Uyarı", f"'{self.weights_file}' dosyasında geçersiz görev adı: '{assignment_name}'. Bu satır atlandı.")

            for assignment in self.assignments:
                if assignment not in self.loaded_weights:
                    messagebox.showwarning("Uyarı", f"'{self.weights_file}' dosyasında '{assignment}' için ağırlık bilgisi bulunamadı. Varsayılan değer (20) kullanılacak.")
                    self.loaded_weights[assignment] = 20

        except FileNotFoundError:
            messagebox.showwarning("Uyarı", f"'{self.weights_file}' dosyası bulunamadı. Varsayılan ağırlıklar kullanılacak.")
            for assignment in self.assignments:
                self.loaded_weights[assignment] = 20
        except Exception as e:
            messagebox.showerror("Hata", f"Ağırlıklar yüklenirken bir hata oluştu: {e}")
            for assignment in self.assignments:
                self.loaded_weights[assignment] = 20

    def save_weights_to_excel(self):
        try:
            weights_data = pd.DataFrame({'Assignment': self.assignments, 'Weight': [self.loaded_weights.get(ass, 20) for ass in self.assignments]})
            weights_data.to_excel(self.weights_file, index=False)
        except Exception as e:
            messagebox.showerror("Hata", f"Ağırlıklar kaydedilirken bir hata oluştu: {e}")

    def validate_matrix_value(self, P):
        if P == "":
            return True
        try:
            value = float(P.replace(',', '.'))
            return 0 <= value <= 1
        except ValueError:
            return False

    def validate_weight_input(self, P):
        if P == "":
            return True
        try:
            value = int(P)
            return 0 <= value <= 100
        except ValueError:
            return False

    def on_entry_focus_in(self, entry):
        if entry.get() == "0":
            entry.delete(0, tk.END)

    def load_from_excel(self):
        try:
            filename = f"tablo_2_3.xlsx"
            if not os.path.exists(filename):
                return

            excel_data = pd.read_excel(filename, sheet_name=['Tablo 2', 'Tablo 3'])

            #ilişkiyi oku
            table2_df = excel_data['Tablo 2'].iloc[2:].copy()
            table2_df['Ders Çıktısı'] = table2_df['Ders Çıktısı'].astype(str)
            outputs = self.get_course_outputs()
            for output_no, _ in outputs:
                output_row = table2_df[table2_df['Ders Çıktısı'] == str(output_no)]
                if not output_row.empty:
                    row_data = output_row.iloc[0]
                    for assignment in self.assignments:
                        if assignment in row_data.index and pd.notna(row_data[assignment]):
                            entry = self.relation_entries[(output_no, assignment)]
                            entry.delete(0, tk.END)
                            entry.insert(0, str(row_data[assignment]))

            self.calculate_sum()
            self.update_weighted_matrix()
            self.check_weights_sum()

        except FileNotFoundError:
            messagebox.showwarning("Uyarı", f"{self.ders_kodu}_{self.ders_adi}_tablo_2_3.xlsx dosyası bulunamadı. Varsayılan değerler kullanılacak.")
        except Exception as e:
            messagebox.showwarning("Uyarı", f"Excel dosyası okunurken bir hata oluştu: {e}\nVarsayılan değerler kullanılacak.")

    def get_course_outputs(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT sira_no, aciklama FROM ders_verileri ORDER BY sira_no")
        outputs = cursor.fetchall()
        conn.close()
        return outputs

    def create_weights_row(self):
        ttk.Label(self.weights_frame, text="Ağırlıklar (%)", font=('Arial', 10, 'bold')).grid(row=0, column=0, padx=5, pady=5)

        validate_cmd = (self.master.register(self.validate_weight_input), '%P')
        for idx, assignment in enumerate(self.assignments, 1):
            weight_var = tk.StringVar(value=str(self.loaded_weights.get(assignment, 20)))
            entry = ttk.Entry(self.weights_frame, width=5, textvariable=weight_var, justify='center', validate='key', validatecommand=validate_cmd)
            entry.grid(row=0, column=idx, padx=2, pady=5, sticky="ew")
            self.weight_entries[assignment] = (weight_var, entry)
            entry.bind('<FocusOut>', lambda event, ass=assignment: self.update_weight(ass))
            entry.bind('<Return>', lambda event, ass=assignment: self.update_weight(ass))

        self.weights_sum_label = ttk.Label(self.weights_frame, text="Toplam: 100")
        self.weights_sum_label.grid(row=0, column=len(self.assignments)+1, padx=5, pady=5)

        for i in range(len(self.assignments) + 2):
            self.weights_frame.grid_columnconfigure(i, weight=1)

        self.check_weights_sum()

    def update_weight(self, assignment):
        try:
            weight_str = self.weight_entries[assignment][0].get()
            new_weight = int(weight_str) if weight_str else 0
            self.loaded_weights[assignment] = new_weight
            self.check_weights_sum()
            self.update_weighted_matrix()
        except ValueError:
            messagebox.showerror("Hata", "Lütfen tam sayı giriniz!")
            self.weight_entries[assignment][0].set(str(self.loaded_weights[assignment]))

    def check_weights_sum(self, event=None):
        total = sum(self.loaded_weights.values())
        self.weights_sum_label.config(text=f"Toplam: {total}")
        self.weights_sum_label.config(foreground='black' if round(total) == 100 else 'red')

    def create_matrix(self):
        outputs = self.get_course_outputs()

        ttk.Label(self.matrix_frame, text="TABLO 2", font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=len(self.assignments)+3, pady=10)

        ttk.Label(self.matrix_frame, text="No", font=('Arial', 10, 'bold')).grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        ttk.Label(self.matrix_frame, text="Açıklama", font=('Arial', 10, 'bold')).grid(row=1, column=1, padx=5, pady=5, sticky='ew')

        for idx, assignment in enumerate(self.assignments, 2):
            ttk.Label(self.matrix_frame, text=assignment, font=('Arial', 10, 'bold')).grid(row=1, column=idx, padx=5, pady=5, sticky="ew")

        ttk.Label(self.matrix_frame, text="TOPLAM", font=('Arial', 10, 'bold')).grid(row=1, column=len(self.assignments)+2, padx=5, pady=5, sticky="ew")

        vcmd_matrix = (self.master.register(self.validate_matrix_value), '%P')

        for row_idx, (output_no, aciklama) in enumerate(outputs, 2):
            ttk.Label(self.matrix_frame, text=str(output_no), font=('Arial', 10)).grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            ttk.Label(self.matrix_frame, text=aciklama, font=('Arial', 10), wraplength=400, justify='left').grid(
                row=row_idx, column=1, padx=5, pady=5, sticky='ew')

            for col_idx, assignment in enumerate(self.assignments, 2):
                entry = ttk.Entry(self.matrix_frame, width=5, validate='key', validatecommand=vcmd_matrix)
                entry.insert(0, "0")
                entry.grid(row=row_idx, column=col_idx, padx=2, pady=2, sticky="ew")
                entry.bind('<FocusIn>', lambda e, entry=entry: self.on_entry_focus_in(entry))
                entry.bind('<KeyRelease>', lambda e: (self.calculate_sum(), self.update_weighted_matrix()))
                self.relation_entries[(output_no, assignment)] = entry

            sum_label = ttk.Label(self.matrix_frame, text="0")
            sum_label.grid(row=row_idx, column=len(self.assignments)+2, padx=5, pady=5, sticky="ew")
            self.sum_labels[output_no] = sum_label

        for i in range(len(self.assignments) + 3):
            self.matrix_frame.grid_columnconfigure(i, weight=1)

    def create_weighted_matrix(self):
        outputs = self.get_course_outputs()

        ttk.Label(self.weighted_matrix_frame, text="TABLO 3", font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=len(self.assignments)+3, pady=10)

        ttk.Label(self.weighted_matrix_frame, text="No", font=('Arial', 10, 'bold')).grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        ttk.Label(self.weighted_matrix_frame, text="Açıklama", font=('Arial', 10, 'bold')).grid(row=1, column=1, padx=5, pady=5, sticky='ew')

        for idx, assignment in enumerate(self.assignments, 2):
            ttk.Label(self.weighted_matrix_frame, text=assignment, font=('Arial', 10, 'bold')).grid(row=1, column=idx, padx=5, pady=5, sticky="ew")

        ttk.Label(self.weighted_matrix_frame, text="TOPLAM", font=('Arial', 10, 'bold')).grid(row=1, column=len(self.assignments)+2, padx=5, pady=5, sticky="ew")

        for row_idx, (output_no, aciklama) in enumerate(outputs, 2):
            ttk.Label(self.weighted_matrix_frame, text=str(output_no), font=('Arial', 10)).grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            ttk.Label(self.weighted_matrix_frame, text=aciklama, font=('Arial', 10), wraplength=400, justify='left').grid(
                row=row_idx, column=1, padx=5, pady=5, sticky='ew')

            row_labels = {}
            for col_idx, assignment in enumerate(self.assignments, 2):
                label = ttk.Label(self.weighted_matrix_frame, text="0.0")
                label.grid(row=row_idx, column=col_idx, padx=2, pady=2, sticky="ew")
                row_labels[assignment] = label

            self.weighted_labels[output_no] = row_labels

            sum_label = ttk.Label(self.weighted_matrix_frame, text="0.0")
            sum_label.grid(row=row_idx, column=len(self.assignments)+2, padx=5, pady=5, sticky="ew")
            self.weighted_sum_labels[output_no] = sum_label

        for i in range(len(self.assignments) + 3):
            self.weighted_matrix_frame.grid_columnconfigure(i, weight=1)

    def calculate_sum(self, event=None):
        for output_no, sum_label in self.sum_labels.items():
            total = 0
            for assignment in self.assignments:
                entry = self.relation_entries[(output_no, assignment)]
                try:
                    value = float(entry.get().replace(',', '.')) if entry.get() else 0
                    total += value
                except ValueError:
                    entry.delete(0, tk.END)
                    entry.insert(0, "0")
            sum_label.config(text=f"{total:.2f}")

    def update_weighted_matrix(self, event=None):
        for output_no in self.weighted_labels:
            weighted_sum = 0
            for assignment in self.assignments:
                try:
                    weight = self.loaded_weights.get(assignment, 0) / 100
                    value = float(self.relation_entries[(output_no, assignment)].get().replace(',', '.'))
                    weighted_value = weight * value
                    self.weighted_labels[output_no][assignment].config(text=f"{weighted_value:.2f}")
                    weighted_sum += weighted_value
                except (ValueError, AttributeError):
                    for label in self.weighted_labels[output_no].values():
                        label.config(text="0.00")
            self.weighted_sum_labels[output_no].config(text=f"{weighted_sum:.2f}")

    def save_data(self):
        if sum(self.loaded_weights.values()) != 100:
            messagebox.showerror("Hata", "Ağırlıkların toplamı 100 olmalıdır!")
            return

        self.save_to_excel()
        self.save_weights_to_excel()

    def save_to_excel(self):
        try:
            outputs = self.get_course_outputs()

            weights_data = {'Ders Çıktısı': '', 'Açıklama': 'Ağırlıklar (%)'}
            for assignment in self.assignments:
                weights_data[assignment] = self.loaded_weights.get(assignment, 0)
            weights_data['TOPLAM'] = sum(self.loaded_weights.values())

            #tablo 2 veri
            data_table2 = []
            for output_no, aciklama in outputs:
                row_data = {'Ders Çıktısı': output_no, 'Açıklama': aciklama}
                for assignment in self.assignments:
                    entry = self.relation_entries[(output_no, assignment)]
                    row_data[assignment] = float(entry.get().replace(',', '.')) if entry.get() else 0
                row_data['TOPLAM'] = float(self.sum_labels[output_no].cget('text'))
                data_table2.append(row_data)

            #tablo 3 veri
            data_table3 = []
            for output_no, aciklama in outputs:
                row_data = {'Ders Çıktısı': output_no, 'Açıklama': aciklama}
                for assignment in self.assignments:
                    label = self.weighted_labels[output_no][assignment]
                    row_data[assignment] = float(label.cget('text'))
                row_data['TOPLAM'] = float(self.weighted_sum_labels[output_no].cget('text'))
                data_table3.append(row_data)

            #excele kaydet
            filename = f"{self.ders_kodu}_{self.ders_adi}_tablo_2_3.xlsx"
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                weights_df_save = pd.DataFrame({'Assignment': self.assignments, 'Weight': [self.loaded_weights.get(ass, 20) for ass in self.assignments]})
                weights_df_save.to_excel(writer, sheet_name='Weights', index=False) 

                pd.DataFrame([weights_data]).to_excel(writer, sheet_name='Tablo 2', index=False, startrow=0) 
                pd.DataFrame(data_table2).to_excel(writer, sheet_name='Tablo 2', startrow=2, index=False)
                pd.DataFrame(data_table3).to_excel(writer, sheet_name='Tablo 3', index=False)

            messagebox.showinfo("Başarılı", f"Veriler {filename} dosyasına kaydedildi!")

        except Exception as e:
            messagebox.showerror("Hata", f"Kaydetme hatası: {str(e)}")

def run_tablo_2_3_app(master, db_path, ders_kodu, ders_adi):
    tablo23_window = tk.Toplevel(master)
    tablo23_window.title(f"{ders_kodu} - {ders_adi} Ders Çıktısı Değerlendirme Matrisi")
    tablo23_window.geometry("1400x800")

    #renkler
    bg_color = "#f0f0f0"
    tablo23_window.configure(bg=bg_color)

    app = CourseOutputMatrixApp(tablo23_window, db_path, ders_kodu, ders_adi)

if __name__ == "__main__":
    pass