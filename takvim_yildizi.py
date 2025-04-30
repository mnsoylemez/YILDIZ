import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import holidays
import pandas as pd

# --- Tatilleri getiren fonksiyon ---
def fetch_holidays():
    result = []
    
    try:
        # Tarihleri kontrol et
        start_str = start_date_entry.get()
        end_str = end_date_entry.get()
        if not start_str or not end_str:
            messagebox.showerror("Hata", "Lütfen başlangıç ve bitiş tarihlerini girin.")
            return

        start_date = datetime.strptime(start_str, "%d/%m/%Y")
        end_date = datetime.strptime(end_str, "%d/%m/%Y")
        if start_date > end_date:
            raise ValueError("Başlangıç tarihi bitiş tarihinden büyük olamaz.")

        # Seçilen ülkeleri al
        selected_countries = []
        for dropdown in [country1_dropdown, country2_dropdown, country3_dropdown]:
            country = dropdown.get()
            if country:
                selected_countries.append(country)

        if not selected_countries:
            messagebox.showerror("Hata", "En az bir ülke seçmelisiniz.")
            return

        date_range = pd.date_range(start=start_date, end=end_date)
        years = {d.year for d in date_range}

        # Tatil verilerini önceden al
        cache = {}
        for code in selected_countries:
            for year in years:
                try:
                    cache[(code, year)] = holidays.country_holidays(code, years=year)
                except Exception:
                    continue

        for day in date_range:
            for code in selected_countries:
                holiday_cal = cache.get((code, day.year))
                if holiday_cal and day in holiday_cal:
                    result.append({
                        "Ülke": code,
                        "Tarih": day.strftime("%d/%m/%Y"),
                        "Tatil": holiday_cal[day]
                    })

        if result:
            export_to_excel(result, start_date, end_date)
        else:
            messagebox.showinfo("Sonuç", "Belirtilen tarihlerde tatil bulunamadı.")

    except ValueError as e:
        messagebox.showerror("Tarih Hatası", str(e))
    except Exception as e:
        messagebox.showerror("Beklenmeyen Hata", str(e))


# --- Excel dosyasına aktarma ---
def export_to_excel(result_data, start_date, end_date):
    df = pd.DataFrame(result_data)
    filename = f"Tatiller_{start_date.strftime('%d-%m-%Y')}_{end_date.strftime('%d-%m-%Y')}.xlsx"
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile=filename,
        filetypes=[("Excel Dosyası", "*.xlsx")]
    )
    if save_path:
        df.to_excel(save_path, index=False)
        messagebox.showinfo("Başarılı", f"Tatiller başarıyla kaydedildi:\n{save_path}")


# --- Arayüz Kurulumu ---
root = tk.Tk()
root.title("Takvim Yıldızı")
root.geometry("480x400")

style = ttk.Style(root)
style.configure("TLabel", font=("Segoe UI", 10))
style.configure("TButton", font=("Segoe UI", 10))
style.configure("TCombobox", font=("Segoe UI", 10))

frame = ttk.Frame(root, padding=15)
frame.pack(fill="both", expand=True)

# Tarih giriş çerçevesi
date_frame = ttk.LabelFrame(frame, text="Tarih Aralığı Seçin", padding=10)
date_frame.pack(fill="x", pady=10)

ttk.Label(date_frame, text="Başlangıç (gg/aa/yyyy):").grid(row=0, column=0, sticky="w")
start_date_entry = ttk.Entry(date_frame)
start_date_entry.grid(row=0, column=1, padx=5, pady=3)

ttk.Label(date_frame, text="Bitiş (gg/aa/yyyy):").grid(row=1, column=0, sticky="w")
end_date_entry = ttk.Entry(date_frame)
end_date_entry.grid(row=1, column=1, padx=5, pady=3)

# Ülke seçim çerçevesi
country_frame = ttk.LabelFrame(frame, text="Ülkeleri Seçin (1-3)", padding=10)
country_frame.pack(fill="x", pady=10)

supported = sorted(list(holidays.list_supported_countries().keys()))

ttk.Label(country_frame, text="Ülke 1:").grid(row=0, column=0, sticky="w")
country1_dropdown = ttk.Combobox(country_frame, values=supported, state="readonly")
country1_dropdown.grid(row=0, column=1, padx=5, pady=2)

ttk.Label(country_frame, text="Ülke 2:").grid(row=1, column=0, sticky="w")
country2_dropdown = ttk.Combobox(country_frame, values=supported, state="readonly")
country2_dropdown.grid(row=1, column=1, padx=5, pady=2)

ttk.Label(country_frame, text="Ülke 3:").grid(row=2, column=0, sticky="w")
country3_dropdown = ttk.Combobox(country_frame, values=supported, state="readonly")
country3_dropdown.grid(row=2, column=1, padx=5, pady=2)

# Buton
ttk.Button(frame, text="Tatilleri Bul ve Kaydet", command=fetch_holidays).pack(pady=20, ipadx=10, ipady=4)

root.mainloop()