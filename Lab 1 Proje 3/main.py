#MAIN
import os
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import tablo_1
import tablo_2_3
import tablo_4
import tablo_5

CRAWLER_DB_NAME = "crawler_data.db" 

def run_script(script_name, ders_kodu, ders_adi):
    try:
        subprocess.run(["python", script_name, ders_kodu, ders_adi], check=True)

        #tablo numarasını script adına göre belirleme
        if script_name == "tablo_1.py":
            tablo_no = "1"
        elif script_name == "tablo_2_3.py":
            tablo_no = "2_3"
        elif script_name == "tablo_4.py":
            tablo_no = "4"
        elif script_name == "tablo_5.py":
            tablo_no = "5"
        elif script_name == "crawler.py":
            tablo_no = "crawler"
        else:
            tablo_no = "bilinmiyor"

        #yeni dosya adını oluşturma
        yeni_dosya_adi = f"{ders_kodu}_{tablo_no}.xlsx"
        eski_dosya_adi = "output.xlsx"  #varsayılan çıktı dosya adı

        #dosyanın var olup olmadığını kontrol et ve yeniden adlandır
        if os.path.exists(eski_dosya_adi):
            try:
                os.rename(eski_dosya_adi, yeni_dosya_adi)
                messagebox.showinfo("Başarılı", f"{script_name} başarıyla çalıştırıldı ve çıktı dosyası {yeni_dosya_adi} olarak kaydedildi.\nDers: {ders_kodu} - {ders_adi}")
            except OSError as e:
                messagebox.showerror("Hata", f"Çıktı dosyası yeniden adlandırılamadı: {e}")
        else:
            messagebox.showerror("Hata", f"{script_name} çalıştırıldı ancak çıktı dosyası bulunamadı.\nDers: {ders_kodu} - {ders_adi}")

    except subprocess.CalledProcessError as e:
        messagebox.showerror("Hata", f"{script_name} çalıştırılamadı. Lütfen kontrol edin.")
    except FileNotFoundError:
        messagebox.showerror("Hata", f"{script_name} dosyası bulunamadı.")

def run_crawler_gui():
    try:
        subprocess.run(["python", "crawler.py"], check=True)
        if os.path.exists(CRAWLER_DB_NAME):
            messagebox.showinfo("Başarılı", f"Veri çekme işlemi tamamlandı ve veritabanı '{CRAWLER_DB_NAME}' oluşturuldu.")
        else:
            messagebox.showinfo("Bilgi", "Veri çekme işlemi tamamlandı.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Hata", f"Veri çekme işlemi sırasında bir hata oluştu. Lütfen kontrol edin.")
    except FileNotFoundError:
        messagebox.showerror("Hata", "crawler.py dosyası bulunamadı.")

def run_tablo_not_gui(ders_kodu, ders_adi):
    try:
        subprocess.run(["python", "tablo_not.py", ders_kodu], check=True)
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Hata", f"Öğrenci notları tablosu açılamadı. Lütfen kontrol edin.")
    except FileNotFoundError:
        messagebox.showerror("Hata", "tablo_not.py dosyası bulunamadı.")

def run_tablo_1_gui(ders_kodu, ders_adi):
    #veritabanı dosyalarının yollarını ana uygulamadaki konuma göre belirleme
    current_dir = os.path.dirname(os.path.abspath(__file__))
    db1_path = os.path.join(current_dir, "ders_ciktilari.db")
    db2_path = os.path.join(current_dir, "program_ciktilari.db")

    #tablo_1 uygulamasını çalıştır
    #main_window argüman olarak geçiriyoruz
    tablo_1.run_tablo_1_app(main_window, db1_path, db2_path, ders_kodu, ders_adi)

def run_tablo_2_3_gui(ders_kodu, ders_adi):
    current_dir = os.path.dirname(os.path.abspath(__file__))
    db_path = os.path.join(current_dir, "ders_ciktilari.db")

    #tablo_2_3 uygulamasını çalıştır
    tablo_2_3.run_tablo_2_3_app(main_window, db_path, ders_kodu, ders_adi)

def run_tablo_4_gui(ders_kodu, ders_adi):
    #doğrudan tablo_4 uygulamasını başlat
    tablo_4.run_tablo_4_gui(main_window, ders_kodu, ders_adi)

def run_tablo_5_gui(ders_kodu, ders_adi):
    tablo_5.run_tablo_5_app(main_window, ders_kodu, ders_adi)

def main():
    print("Ana program başlatılıyor...")

    global main_window  #main_window'u global olarak tanımla
    main_window = tk.Tk()
    main_window.title("Program Çıktıları İlişki Matrisi Uygulaması")
    main_window.geometry("700x500")
    main_window.configure(bg="#FFB6C1")

    #stil
    style = ttk.Style(main_window)
    style.theme_use('clam')

    #renkler
    bg_color = "#FFB6C1"
    primary_color = "#dd91b9"
    secondary_color = "#93d2e3"
    accent_color = "#d7b3f0"

    main_window.configure(bg=bg_color)

    style.configure('TFrame', background=bg_color)
    style.configure('TLabelframe', background=bg_color, borderwidth=2, relief="groove")
    style.configure('TLabelframe.Label', foreground='black', font=('Arial', 12, 'bold'))
    style.configure('TLabel', background=bg_color, foreground='black')
    style.configure('TButton', background=secondary_color, foreground='black', font=('Arial', 10))
    style.map('TButton', background=[('active', accent_color)])
    style.configure('TEntry', fieldbackground='white', foreground='black')
    style.configure('TListbox', background='white', foreground='black')

    main_frame = ttk.Frame(main_window, padding="20")
    main_frame.pack(expand=True, fill="both")

    title_label = ttk.Label(main_frame, text="Ders Seçimi ve İşlemler", font=("Arial", 18, "bold"))
    title_label.pack(pady=20)

    #ders ekleme
    ders_ekleme_frame = ttk.LabelFrame(main_frame, text="Ders Ekle", padding=10)
    ders_ekleme_frame.pack(pady=10, fill="both", expand=True)
    ders_ekleme_frame.grid_columnconfigure(1, weight=1)
    ders_ekleme_frame.grid_columnconfigure(2, weight=1)

    ders_kodu_label = ttk.Label(ders_ekleme_frame, text="Ders Kodu:")
    ders_kodu_label.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
    ders_kodu_entry = ttk.Entry(ders_ekleme_frame)
    ders_kodu_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

    ders_adi_label = ttk.Label(ders_ekleme_frame, text="Ders Adı:")
    ders_adi_label.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
    ders_adi_entry = ttk.Entry(ders_ekleme_frame)
    ders_adi_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

    DERSLER_DOSYASI = "kayitli_dersler.txt"
    dersler = []

    #kaıtlı dersleri yükleme
    if os.path.exists(DERSLER_DOSYASI):
        with open(DERSLER_DOSYASI, "r", encoding="utf-8") as f:
            for line in f:
                dersler.append(line.strip())

    ders_listesi_var = tk.StringVar(value=dersler)
    ders_listesi = tk.Listbox(ders_ekleme_frame, listvariable=ders_listesi_var, height=5)
    ders_listesi.grid(row=0, column=2, rowspan=2, padx=10, pady=5, sticky="nsew")

    def ders_ekle():
        kodu = ders_kodu_entry.get().strip()
        adi = ders_adi_entry.get().strip()
        if kodu and adi:
            ders = f"{kodu} - {adi}"
            if ders not in dersler:
                dersler.append(ders)
                ders_listesi_var.set(dersler)
            ders_kodu_entry.delete(0, tk.END)
            ders_adi_entry.delete(0, tk.END)
        else:
            messagebox.showerror("Hata", "Lütfen ders kodu ve adını girin.")

    def ders_sil():
        secili_ders_index = ders_listesi.curselection()
        if secili_ders_index:
            secili_ders = dersler[secili_ders_index[0]]
            dersler.remove(secili_ders)
            ders_listesi_var.set(dersler)
        else:
            messagebox.showerror("Hata", "Lütfen silmek için bir ders seçin.")

    ekle_button = ttk.Button(ders_ekleme_frame, text="Ders Ekle", command=ders_ekle)
    ekle_button.grid(row=2, column=0, columnspan=2, pady=5, sticky="ew")

    sil_button = ttk.Button(ders_ekleme_frame, text="Seçili Dersi Sil", command=ders_sil)
    sil_button.grid(row=2, column=2, pady=5, sticky="ew")

    ders_ekleme_frame.grid_rowconfigure(2, weight=0)

    def calistir_secili_ders(func):
        secili_ders_index = ders_listesi.curselection()
        if secili_ders_index:
            secili_ders = dersler[secili_ders_index[0]]
            ders_kodu, ders_adi = secili_ders.split(" - ", 1)
            func(ders_kodu, ders_adi)
        else:
            messagebox.showerror("Hata", "Lütfen bir ders seçin.")

    #işlem butonları
    islemler_frame = ttk.LabelFrame(main_frame, text="İşlemler", padding=10)
    islemler_frame.pack(pady=10, fill="both", expand=True)
    islemler_frame.grid_columnconfigure(0, weight=1, uniform="group1")
    islemler_frame.grid_columnconfigure(1, weight=1, uniform="group1")

    crawler_button = ttk.Button(islemler_frame, text="Veri Çekme (Crawler)", command=run_crawler_gui)
    crawler_button.grid(row=0, column=0, pady=5, padx=5, sticky="ew")

    tablo_not_button = ttk.Button(islemler_frame, text="Öğrenci Notları Tablosu Düzenleme", command=lambda: calistir_secili_ders(run_tablo_not_gui))
    tablo_not_button.grid(row=0, column=1, pady=5, padx=5, sticky="ew")

    tablo_1_button = ttk.Button(islemler_frame, text="Program-Ders Çıktısı İlişki Matrisi (Tablo 1)", command=lambda: calistir_secili_ders(run_tablo_1_gui))
    tablo_1_button.grid(row=1, column=0, pady=5, padx=5, sticky="ew")

    tablo_2_3_button = ttk.Button(islemler_frame, text="Ders Çıktısı Değerlendirme Matrisi (Tablo 2-3)", command=lambda: calistir_secili_ders(run_tablo_2_3_gui))
    tablo_2_3_button.grid(row=1, column=1, pady=5, padx=5, sticky="ew")

    tablo_4_button = ttk.Button(islemler_frame, text="Öğrenci Ders Çıktıları Hesaplama (Tablo 4)", command=lambda: calistir_secili_ders(run_tablo_4_gui))
    tablo_4_button.grid(row=2, column=0, pady=5, padx=5, sticky="ew")

    tablo_5_button = ttk.Button(islemler_frame, text="Program Çıktısı Başarı Oranları (Tablo 5)", command=lambda: calistir_secili_ders(run_tablo_5_gui))
    tablo_5_button.grid(row=2, column=1, pady=5, padx=5, sticky="ew")

    main_frame.grid_rowconfigure(2, weight=1)
    main_frame.grid_columnconfigure(0, weight=1)

    def on_closing():
        #dersleri kaydetme
        with open(DERSLER_DOSYASI, "w", encoding="utf-8") as f:
            for ders in dersler:
                f.write(ders + "\n")
        main_window.destroy()

    main_window.protocol("WM_DELETE_WINDOW", on_closing)
    main_window.mainloop()

if __name__ == "__main__":
    main()