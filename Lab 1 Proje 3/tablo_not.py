import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
import sys
import os

class ProgramMatrisUygulamasi:
    def __init__(self, root, ders_kodu):
        self.root = root
        self.root.title("Öğrenci Notları Düzenleme")
        self.root.geometry("1200x800")
        self.ders_kodu = ders_kodu
        self.excel_file = 'ogrenci_notlari.xlsx'
        self.weights_file = f'{ders_kodu}_weights.xlsx'
        self.df = None
        self.weights = {}
        self.weight_entries = {}
        self.tree = None

        self.load_data()
        self.create_widgets()

    def load_data(self):
        try:
            self.df = pd.read_excel(self.excel_file)
            self.load_weights()
            if 'Ortalama' not in self.df.columns:
                self.df['Ortalama'] = 0.0
            self.df['Ortalama'] = self.calculate_weighted_average(self.df)
        except FileNotFoundError:
            self.df = pd.DataFrame()
            messagebox.showwarning("Uyarı", f"{self.excel_file} dosyası bulunamadı. Boş bir tablo oluşturuldu.")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel dosyası yüklenirken hata oluştu: {str(e)}")
            self.df = pd.DataFrame()

    def load_weights(self):
        try:
            weights_df = pd.read_excel(self.weights_file)
            for col in self.df.columns[1:]:
                if col != "Ortalama":
                    try:
                        self.weights[col] = int(weights_df[weights_df["Assignment"] == col]["Weight"].iloc[0])
                    except IndexError:
                        messagebox.showerror("Hata", f"{self.weights_file} dosyası '{col}' ödevini içermiyor.")
                        self.weights[col] = 0
        except FileNotFoundError:
            num_assignments = len([col for col in self.df.columns[1:] if col != "Ortalama"])
            if num_assignments > 0:
                default_weight = 100 // num_assignments
                remainder = 100 % num_assignments
                for i, col in enumerate(self.df.columns[1:]):
                    if col != "Ortalama":
                        self.weights[col] = default_weight + (1 if i < remainder else 0)

    def validate_weight_input(self, P):
        if P == "":
            return True
        try:
            value = int(P)
            return 0 <= value <= 100
        except ValueError:
            return False

    def update_weight(self, column):
        try:
            weight_str = self.weight_entries[column][0].get()
            new_weight = int(weight_str) if weight_str else 0
            self.weights[column] = new_weight

            #tüm ağırlıkları topla
            total_weight = sum(int(self.weight_entries[col][0].get() or 0) for col in self.weights.keys())

            if total_weight == 100:
                #ağırlıklar toplamı 100 ise ortalamaları güncelle
                self.df['Ortalama'] = self.calculate_weighted_average(self.df)
                self.load_data_to_treeview()
            else:
                messagebox.showwarning("Uyarı", f"Ağırlıkların toplamı 100 olmalıdır! Şu anki toplam: {total_weight}")

        except ValueError:
            messagebox.showerror("Hata", "Lütfen tam sayı giriniz!")
            self.weight_entries[column][0].set(str(self.weights[column]))
            return

    def create_widgets(self):
        weights_frame = ttk.Frame(self.root)
        weights_frame.grid(row=0, column=0, columnspan=2, sticky='ew', padx=10, pady=5)

        padx_val = 5
        pady_val = 2
        col_num = 1

        for col in self.df.columns[1:]:
            if col != "Ortalama":
                ttk.Label(weights_frame, text=f"{col} (%):").grid(row=0, column=col_num-1, padx=padx_val, pady=pady_val)
                weight_var = tk.StringVar(value=str(self.weights.get(col, 0)))
                entry = ttk.Entry(weights_frame, width=5, textvariable=weight_var, justify='center')
                entry.grid(row=1, column=col_num-1, padx=padx_val, pady=pady_val)
                self.weight_entries[col] = (weight_var, entry)

                validate_cmd = (self.root.register(self.validate_weight_input), '%P')
                entry.config(validate="key", validatecommand=validate_cmd)

                entry.bind('<FocusOut>', lambda event, c=col: self.update_weight(c))
                entry.bind('<Return>', lambda event, c=col: self.update_weight(c))
                col_num += 1

        self.tree = ttk.Treeview(self.root, columns=list(self.df.columns), show='headings')
        self.tree.grid(row=2, column=0, columnspan=2, sticky='nsew')

        for col in self.df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor='center')

        self.load_data_to_treeview()

        def update_cell(event, item, column):
            if not item:
                return

            entry = event.widget
            col_index = int(column.replace('#', '')) - 1
            row_id = self.tree.index(item)
            col_name = self.df.columns[col_index]
            new_value = entry.get()

            if col_name == self.df.columns[0]:
                self.df.iloc[row_id, col_index] = str(new_value)
            else:
                try:
                    new_value = float(new_value)
                    if col_name != self.df.columns[0] and 0 <= new_value <= 100:
                        self.df.iloc[row_id, col_index] = new_value
                        self.df.loc[row_id, 'Ortalama'] = self.calculate_weighted_average(self.df.iloc[[row_id]])[0]
                    elif col_name != self.df.columns[0]:
                        messagebox.showerror("Hata", "Not değeri 0 ile 100 arasında olmalıdır.")
                except ValueError:
                    messagebox.showerror("Hata", "Geçersiz değer")
            self.load_data_to_treeview()

        def on_double_click(event):
            item = self.tree.identify_row(event.y)
            column = self.tree.identify_column(event.x)

            if not item:
                return

            col_index = int(column.replace('#', '')) - 1
            row_id = self.tree.index(item)
            cell_value = self.df.iloc[row_id, col_index]

            entry = ttk.Entry(self.root)
            entry.insert(0, cell_value)
            entry.place(x=event.x_root - self.root.winfo_rootx(),
                      y=event.y_root - self.root.winfo_rooty())
            entry.focus_set()

            entry.bind('<Return>', lambda e: [update_cell(e, item, column), entry.destroy()])
            entry.bind('<FocusOut>', lambda e: entry.destroy())

        self.tree.bind('<Double-1>', on_double_click)

        edit_frame = ttk.Frame(self.root)
        edit_frame.grid(row=3, column=0, columnspan=2, pady=10)

        ttk.Button(edit_frame, text="Kaydet",
                  command=self.save_excel).grid(row=0, column=0, padx=5)

        def add_row():
            new_row = {col: '' if col == self.df.columns[0] else 0.0 for col in self.df.columns}
            self.df.loc[len(self.df)] = new_row
            self.load_data_to_treeview()

        def delete_row():
            selected_item = self.tree.focus()
            if selected_item:
                row_id = self.tree.index(selected_item)
                self.df.drop(self.df.index[row_id], inplace=True)
                self.df.reset_index(drop=True, inplace=True)
                self.load_data_to_treeview()

        ttk.Button(edit_frame, text="Satır Ekle",
                  command=add_row).grid(row=0, column=1, padx=5)
        ttk.Button(edit_frame, text="Satır Sil",
                  command=delete_row).grid(row=0, column=2, padx=5)

        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(2, weight=1)

    def load_data_to_treeview(self):
        self.tree.delete(*self.tree.get_children())
        for _, row in self.df.iterrows():
            row_values = [str(row[col]) if col == self.df.columns[0] else
                         (f"{row['Ortalama']:.1f}" if col == 'Ortalama' else f"{row[col]:.1f}")
                         for col in self.df.columns]
            self.tree.insert('', 'end', values=row_values)

    def save_excel(self):
        total_weight = sum(int(self.weight_entries[col][0].get() or 0) for col in self.weights.keys())

        if total_weight != 100:
            messagebox.showerror("Hata", f"Ödev ağırlıklarının toplamı 100 olmalıdır! Şu anki toplam: {total_weight}")
            return

        #ağırlıkları güncelle ve kaydet
        for col in self.weights.keys():
            try:
                self.weights[col] = int(self.weight_entries[col][0].get() or 0)
            except ValueError:
                messagebox.showerror("Hata", f"{col} için geçersiz ağırlık değeri!")
                return

        #ortalamaları yeniden hesapla
        self.df['Ortalama'] = self.calculate_weighted_average(self.df)

        new_file_name = f"{self.ders_kodu}_not.xlsx"

        try:
            self.df.to_excel(new_file_name, index=False)
            self.save_weights()
            messagebox.showinfo("Başarılı", f"Veriler {new_file_name} dosyasına kaydedildi.")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel'e kaydetme hatası: {str(e)}")

    def save_weights(self):
        try:
            weights_data = []
            for assignment, weight in self.weights.items():
                weights_data.append({"Assignment": assignment, "Weight": weight})
            weights_df = pd.DataFrame(weights_data)
            weights_df.to_excel(self.weights_file, index=False)
        except Exception as e:
            messagebox.showerror("Hata", f"Ağırlıklar kaydedilirken hata oluştu: {e}")

    def calculate_weighted_average(self, df):
        weighted_sum = pd.Series(0.0, index=df.index)

        for col in df.columns[1:]:
            if col != "Ortalama" and col in self.weights:
                weight_percentage = self.weights[col] / 100
                weighted_sum += df[col].fillna(0) * weight_percentage

        return weighted_sum

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Kullanım: python tablo_not.py <ders_kodu>")
        sys.exit(1)

    ders_kodu = sys.argv[1]
    root = tk.Tk()
    app = ProgramMatrisUygulamasi(root, ders_kodu)
    root.mainloop()