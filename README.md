# YILDIZ - Holiday Calendar Tools

YILDIZ is a collection of two Python-based applications designed to fetch and manage national holidays across various countries. The project includes two scripts: `yildiz_takvimi.py` (Yıldız Takvimi) and `takvim_yildizi.py` (Takvim Yıldızı). Both tools utilize the `holidays` library to retrieve holiday data and provide user-friendly GUIs built with `tkinter`. Results can be exported to Excel files for easy analysis.

## Features

- **yildiz_takvimi.py (Yıldız Takvimi)**:
  - Fetch holidays for a specific date across all supported countries.
  - Fetch holidays within a date range across all supported countries.
  - Fetch holidays for a specific year and country.
  - Display results in a table with an option to export to Excel.
- **takvim_yildizi.py (Takvim Yıldızı)**:
  - Fetch holidays within a date range for up to three selected countries.
  - Export results directly to an Excel file.
- User-friendly GUIs for both tools.
- Supports exporting holiday data to Excel for further analysis.

## Prerequisites

- Python 3.6 or higher
- Required Python packages (listed in `requirements.txt`)

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/mnsoylemez/YILDIZ.git
   cd YILDIZ
   ```
2. Create a virtual environment (optional but recommended):

   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
3. Install the dependencies:

   ```bash
   pip install -r requirements.txt
   ```
4. Run either application:
   - For Yıldız Takvimi:

     ```bash
     python yildiz_takvimi.py
     ```
   - For Takvim Yıldızı:

     ```bash
     python takvim_yildizi.py
     ```

## Usage

### Using `yildiz_takvimi.py` (Yıldız Takvimi)

1. Launch the application:

   ```bash
   python yildiz_takvimi.py
   ```
2. Choose a query type:
   - **Belirli Tarih (Specific Date)**: Enter a date (e.g., `01/01/2024`) to find holidays on that day across all supported countries.
   - **Zaman Aralığı (Date Range)**: Enter a start and end date (e.g., `01/01/2024` to `31/12/2024`) to find holidays in that range across all supported countries.
   - **Belirli Yıl ve Ülke (Specific Year and Country)**: Enter a year (e.g., `2024`) and select a country from the dropdown to find holidays for that year.
3. Click **"Tatilleri Bul"** to fetch the holidays.
4. View the results in a pop-up window and click **"Excel'e Aktar"** to export the data to an Excel file.

### Using `takvim_yildizi.py` (Takvim Yıldızı)

1. Launch the application:

   ```bash
   python takvim_yildizi.py
   ```
2. Enter the date range:
   - **Başlangıç (Start Date)**: Enter the start date (e.g., `01/01/2024`).
   - **Bitiş (End Date)**: Enter the end date (e.g., `31/12/2024`).
3. Select up to three countries from the dropdown menus (e.g., `Turkey`, `UnitedStates`, `Germany`).
4. Click **"Tatilleri Bul ve Kaydet"** to fetch the holidays and export them to an Excel file.
5. Choose a file path to save the Excel file when prompted.

## Output

- **yildiz_takvimi.py**: Results are displayed in a table with columns `Holiday`, `Country`, and `Date`. The Excel export filename is automatically suggested based on the query (e.g., `Ulusal Tatiller 01-01-2024.xlsx`).
- **takvim_yildizi.py**: Results are exported directly to an Excel file with columns `Ülke` (Country), `Tari f` (Date), and `Tatil` (Holiday). The filename is formatted as `Tatiller_[start_date]_[end_date].xlsx` (e.g., `Tatiller_01-01-2024_31-12-2024.xlsx`).

## Notes

- Dates must be entered in the format `dd/mm/yyyy`.
- The applications use the `holidays` library, which supports a wide range of countries. However, some countries may not have holiday data available for certain years.
- Ensure you have write permissions in the directory where you save the Excel files.

## Dependencies

See `requirements.txt` for the full list of dependencies. Key libraries include:

- `tkinter` (for the GUI)
- `pandas` (for data handling and Excel export)
- `holidays` (for fetching national holidays)
- `openpyxl` (for Excel file handling)

## Contributing

Contributions are welcome! If you have ideas for new features (e.g., support for custom holidays, additional query options), please open an issue or submit a pull request on GitHub.

## Contact

For questions or feedback, please contact soylemeznurhan@gmail.com.
