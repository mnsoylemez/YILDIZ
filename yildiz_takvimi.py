import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import holidays
import pandas as pd

# --- Holiday Fetching Logic ---
def fetch_holidays():
    result = []
    filename = "Ulusal Tatiller.xlsx"
    supported_countries_with_names = holidays.list_supported_countries(True)
    supported_countries_dict = holidays.list_supported_countries()  # CODE: [provinces]

    try:
        if specific_date_var.get():
            try:
                date_str = date_entry.get()
                if not date_str:
                    raise ValueError("Tarih alanı boş olamaz.")
                date = datetime.strptime(date_str, "%d/%m/%Y")
            except ValueError as e:
                messagebox.showerror("Girdi Hatası", f"Geçersiz tarih formatı: {e}")
                return

            for country_code in supported_countries_dict.keys():
                try:
                    country_holidays = holidays.country_holidays(country_code, years=date.year)
                    if date in country_holidays:
                        country_name = country_code
                        for name, codes in supported_countries_with_names.items():
                            if country_code in codes:
                                country_name = name
                                break
                        result.append({
                            "Holiday": country_holidays[date],
                            "Country": country_name,
                            "Date": date.strftime("%d/%m/%Y")
                        })
                except Exception as e:
                    print(f"{country_code} kodu için hata: {e}")

            filename = f"Ulusal Tatiller {date.strftime('%d-%m-%Y')}.xlsx"

        elif specific_period_var.get():
            try:
                start_date = datetime.strptime(start_date_entry.get(), "%d/%m/%Y")
                end_date = datetime.strptime(end_date_entry.get(), "%d/%m/%Y")
                if start_date > end_date:
                    raise ValueError("Başlangıç tarihi bitiş tarihinden büyük olamaz.")
            except ValueError as e:
                messagebox.showerror("Girdi Hatası", f"Geçersiz tarih aralığı: {e}")
                return

            date_range = pd.date_range(start=start_date, end=end_date)
            years = {d.year for d in date_range}

            cache = {}
            for country_code in supported_countries_dict.keys():
                for year in years:
                    try:
                        cache[(country_code, year)] = holidays.country_holidays(country_code, years=year)
                    except Exception:
                        continue

            for single_date in date_range:
                for country_code in supported_countries_dict.keys():
                    holiday_cal = cache.get((country_code, single_date.year))
                    if holiday_cal and single_date in holiday_cal:
                        country_name = country_code
                        for name, codes in supported_countries_with_names.items():
                            if country_code in codes:
                                country_name = name
                                break
                        result.append({
                            "Holiday": holiday_cal[single_date],
                            "Country": country_name,
                            "Date": single_date.strftime("%d/%m/%Y")
                        })

            filename = f"Ulusal Tatiller {start_date.strftime('%d-%m-%Y')} ile {end_date.strftime('%d-%m-%Y')}.xlsx"

        elif specific_year_var.get():
            try:
                year = int(year_entry.get())
            except ValueError:
                messagebox.showerror("Girdi Hatası", "Geçerli bir yıl girin (örn: 2024).")
                return

            selected_country_name = country_dropdown.get()
            if not selected_country_name:
                messagebox.showerror("Girdi Hatası", "Bir ülke seçin.")
                return

            # Find the correct country code for the selected country name
            country_code = None
            for code, country_names in supported_countries_with_names.items():
                if selected_country_name == code:
                    country_code = country_names[0] if country_names else code
                    break
                
            if not country_code:
                # Try reverse lookup - find which country code contains this name
                for code, names in supported_countries_dict.items():
                    for name, codes in supported_countries_with_names.items():
                        if code in codes and name == selected_country_name:
                            country_code = code
                            break
                    if country_code:
                        break

            if not country_code:
                messagebox.showerror("Hata", f"'{selected_country_name}' için ülke kodu bulunamadı.")
                return
                
            try:
                print(f"Trying to get holidays for {selected_country_name} with code {country_code} for year {year}")
                
                # Try the direct country_holidays method first
                try:
                    country_holidays = holidays.country_holidays(country_code, years=year)
                except Exception as e:
                    print(f"First method failed: {e}")
                    # If that fails, try using the country class directly
                    try:
                        country_class = getattr(holidays, country_code)
                        country_holidays = country_class(years=year)
                    except AttributeError as e:
                        print(f"Second method failed: {e}")
                        # Try one more approach - iterating through all country classes
                        found = False
                        for attr_name in dir(holidays):
                            if attr_name.upper() == country_code.upper():
                                country_class = getattr(holidays, attr_name)
                                if callable(country_class):
                                    try:
                                        country_holidays = country_class(years=year)
                                        found = True
                                        break
                                    except Exception:
                                        continue
                        
                        if not found:
                            raise Exception(f"'{selected_country_name}' için tatil verisi bulunamadı.")

            except Exception as e:
                messagebox.showerror("Hata", f"{selected_country_name} için tatil bulunamadı: {str(e)}")
                return

            for date, name in sorted(country_holidays.items()):
                result.append({
                    "Holiday": name,
                    "Date": date.strftime("%d/%m/%Y"),
                    "Country": selected_country_name
                })

            filename = f"{selected_country_name} {year} Ulusal Tatilleri.xlsx"

        if result:
            display_results(result, filename)
        else:
            messagebox.showinfo("Sonuç Yok", "Tatil bulunamadı.")

    except Exception as e:
        messagebox.showerror("Hata", f"Beklenmeyen hata: {str(e)}")
        import traceback
        traceback.print_exc()  # This will print the full error trace to the console

# --- Display ---
def display_results(result_data, suggested_filename):
    window = tk.Toplevel(root)
    window.title("Sonuçlar")
    window.geometry("700x400")

    columns = list(result_data[0].keys())

    frame = ttk.Frame(window)
    frame.pack(expand=True, fill="both")

    tree = ttk.Treeview(frame, columns=columns, show="headings")
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=200)

    for row in result_data:
        values = [row.get(col, "") for col in columns]
        tree.insert("", tk.END, values=values)

    tree.pack(expand=True, fill="both")

    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")

    def export_to_excel():
        df = pd.DataFrame(result_data)
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=suggested_filename,
            filetypes=[("Excel Dosyası", "*.xlsx")]
        )
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Başarılı", f"Kaydedildi:\n{save_path}")

    ttk.Button(window, text="Excel'e Aktar", command=export_to_excel).pack(pady=10)

# --- GUI ---# --- GUI Styling Improvements ---
root = tk.Tk()
root.title("Yıldız Takvimi")
root.geometry("520x480")

style = ttk.Style(root)
style.configure("TLabel", font=("Segoe UI", 10))
style.configure("TButton", font=("Segoe UI", 10))
style.configure("TCheckbutton", font=("Segoe UI", 10))
style.configure("TCombobox", font=("Segoe UI", 10))

specific_date_var = tk.BooleanVar(value=True)
specific_period_var = tk.BooleanVar(value=False)
specific_year_var = tk.BooleanVar(value=False)

def toggle_inputs(selected_var):
    for var in [specific_date_var, specific_period_var, specific_year_var]:
        if var != selected_var:
            var.set(False)

    if not (specific_date_var.get() or specific_period_var.get() or specific_year_var.get()):
        selected_var.set(True)

    date_entry.config(state="normal" if specific_date_var.get() else "disabled")
    start_date_entry.config(state="normal" if specific_period_var.get() else "disabled")
    end_date_entry.config(state="normal" if specific_period_var.get() else "disabled")
    year_entry.config(state="normal" if specific_year_var.get() else "disabled")
    country_dropdown.config(state="readonly" if specific_year_var.get() else "disabled")

frame = ttk.Frame(root, padding=15)
frame.pack(expand=True, fill="both")

query_frame = ttk.LabelFrame(frame, text="Sorgu Tipi", padding=10)
query_frame.pack(fill="x", pady=8)

ttk.Checkbutton(query_frame, text="Belirli Tarih", variable=specific_date_var, command=lambda: toggle_inputs(specific_date_var)).pack(anchor="w", pady=2)
ttk.Checkbutton(query_frame, text="Zaman Aralığı", variable=specific_period_var, command=lambda: toggle_inputs(specific_period_var)).pack(anchor="w", pady=2)
ttk.Checkbutton(query_frame, text="Belirli Yıl ve Ülke", variable=specific_year_var, command=lambda: toggle_inputs(specific_year_var)).pack(anchor="w", pady=2)

input_frame = ttk.LabelFrame(frame, text="Giriş Bilgileri", padding=10)
input_frame.pack(fill="x", pady=8)

def labeled_entry(row, label_text, entry_widget):
    ttk.Label(input_frame, text=label_text).grid(row=row, column=0, sticky="w", pady=3)
    entry_widget.grid(row=row, column=1, sticky="ew", pady=3)

date_entry = ttk.Entry(input_frame, width=25)
start_date_entry = ttk.Entry(input_frame, width=25)
end_date_entry = ttk.Entry(input_frame, width=25)
year_entry = ttk.Entry(input_frame, width=25)

labeled_entry(0, "Tarih (gg/aa/yyyy):", date_entry)
labeled_entry(1, "Başlangıç (gg/aa/yyyy):", start_date_entry)
labeled_entry(2, "Bitiş (gg/aa/yyyy):", end_date_entry)
labeled_entry(3, "Yıl:", year_entry)

country_list = sorted(list(holidays.list_supported_countries(True).keys()))
ttk.Label(input_frame, text="Ülke:").grid(row=4, column=0, sticky="w", pady=3)
country_dropdown = ttk.Combobox(input_frame, values=country_list, state="disabled", width=22)
country_dropdown.grid(row=4, column=1, sticky="ew", pady=3)

input_frame.columnconfigure(1, weight=1)

ttk.Button(frame, text="Tatilleri Bul", command=fetch_holidays).pack(pady=15, ipadx=10, ipady=3)

toggle_inputs(specific_date_var)


root.mainloop()