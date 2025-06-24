import pandas as pd
from fpdf import FPDF
import matplotlib.pyplot as plt
import os
from datetime import datetime

# --- Step 1: Read data from Excel or fallback to CSV ---
file_path_excel = 'sample_data.xlsx'
file_path_csv = 'sample_data.csv'

try:
    if os.path.exists(file_path_excel):
        data = pd.read_excel(file_path_excel)
        print("✅ Loaded data from Excel")
    elif os.path.exists(file_path_csv):
        data = pd.read_csv(file_path_csv)
        print("✅ Loaded data from CSV")
    else:
        raise FileNotFoundError("Neither Excel nor CSV data files found.")
except Exception as e:
    print(f"❌ Failed to load data: {e}")
    exit(1)

# --- Step 2: Basic analysis ---
avg_temp = data['Temperature'].mean()
avg_humidity = data['Humidity'].mean()
avg_wind = data['Wind Speed'].mean()

# --- Step 3: Visualizations ---

# Weather Trends Line Chart
plt.figure(figsize=(10, 5))
plt.plot(data['Datetime'], data['Temperature'], color='orange', marker='o', label='Temperature (°C)')
plt.plot(data['Datetime'], data['Humidity'], color='blue', marker='s', label='Humidity (%)')
plt.plot(data['Datetime'], data['Wind Speed'], color='green', marker='^', label='Wind Speed (m/s)')
plt.xlabel('Datetime')
plt.ylabel('Values')
plt.title('Weather Trends Over Time')
plt.xticks(rotation=45)
plt.legend()
plt.tight_layout()
plt.savefig('weather_trends.png')
plt.close()

# Temperature Range Distribution Pie Chart
temp_bins = pd.cut(data['Temperature'], bins=[-10, 0, 10, 20, 30, 40])
temp_dist = temp_bins.value_counts().sort_index()

plt.figure(figsize=(6, 6))
temp_dist.plot.pie(autopct='%1.1f%%', startangle=140, colors=plt.cm.Pastel1.colors)
plt.title('Temperature Range Distribution')
plt.ylabel('')
plt.tight_layout()
plt.savefig('temp_distribution.png')
plt.close()

# --- Step 4: PDF Report Class ---
class PDFReport(FPDF):
    def header(self):
        logo_path = 'logo.png'
        if os.path.exists(logo_path):
            self.image(logo_path, 10, 8, 33)
        self.set_font('Arial', 'B', 16)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, 'Automated Weather Report', ln=True, align='C')
        self.set_font('Arial', '', 10)
        self.cell(0, 10, f'Report generated on: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', ln=True, align='C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 10)
        self.set_text_color(128)
        self.cell(0, 10, f'Page {self.page_no()}', align='C')

    def section_title(self, title):
        self.set_fill_color(0, 51, 102)
        self.set_text_color(255, 255, 255)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, title, 0, 1, 'L', True)
        self.set_text_color(0, 0, 0)

    def add_summary_table(self, stats_dict):
        self.set_font('Arial', '', 12)
        line_height = self.font_size * 2.5
        col_width = (self.w - 2 * self.l_margin) / 3

        # Table header
        self.set_fill_color(200, 220, 255)
        self.cell(col_width, line_height, "Metric", border=1, fill=True)
        self.cell(col_width, line_height, "Value", border=1, fill=True)
        self.cell(col_width, line_height, "Units", border=1, fill=True)
        self.ln(line_height)

        # Table content
        fill = False
        for key, (value, unit) in stats_dict.items():
            self.set_fill_color(235, 245, 255) if fill else self.set_fill_color(255, 255, 255)
            self.cell(col_width, line_height, key, border=1, fill=True)
            self.cell(col_width, line_height, f"{value:.2f}", border=1, fill=True)
            self.cell(col_width, line_height, unit, border=1, fill=True)
            self.ln(line_height)
            fill = not fill

# --- Step 5: Build the PDF Report ---
pdf = PDFReport()
pdf.add_page()

pdf.section_title('Summary Statistics')
summary_stats = {
    "Average Temperature": (avg_temp, "°C"),
    "Average Humidity": (avg_humidity, "%"),
    "Average Wind Speed": (avg_wind, "m/s"),
}
pdf.add_summary_table(summary_stats)

pdf.ln(10)
pdf.section_title('Visualizations')

# Trends Plot
if os.path.exists('weather_trends.png'):
    pdf.cell(0, 10, 'Weather Trends Over Time:', ln=True)
    pdf.image('weather_trends.png', x=10, w=190)
else:
    pdf.cell(0, 10, 'Weather Trends image not found.', ln=True)

pdf.ln(10)

# Pie Chart
if os.path.exists('temp_distribution.png'):
    pdf.cell(0, 10, 'Temperature Range Distribution:', ln=True)
    pdf.image('temp_distribution.png', x=60, w=90)
else:
    pdf.cell(0, 10, 'Temperature Distribution image not found.', ln=True)

# --- Save the PDF ---
pdf.output('automated_weather_report.pdf')
print("✅ Report generated as 'automated_weather_report.pdf'")
