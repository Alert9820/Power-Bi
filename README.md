# 📊 BI Reporter — Automated Business Intelligence Tool

> Upload any CSV or Excel file. Get back a fully formatted 5-sheet Excel report with KPIs, embedded charts, ML revenue prediction, and Power BI integration — in seconds.

🔗 **Live Demo:** [your-app.onrender.com](https://power-bi-0hii.onrender.com/)

---

## ✨ What It Does

```
Upload CSV / Excel
      ↓
Auto Clean (duplicates · missing values · outliers)
      ↓
Feature Engineering (Profit · Margins · Date parts)
      ↓
Train ML Model (Random Forest → revenue prediction)
      ↓
Generate Excel Report (5 sheets · charts embedded)
      ↓
Download & Connect to Power BI
```

---

## 📁 Features

### 🧹 Auto ETL Pipeline
- Reads CSV and Excel (.xlsx, .xls) files up to 200MB
- Removes duplicate rows
- Imputes missing values — median for numeric, mode for categorical
- Removes outliers using IQR method (1st–99th percentile)
- Auto-engineers columns: `Profit`, `Profit_Margin_%`, `Total_Value`

### 📊 Excel Report (5 Sheets)
| Sheet | Content |
|---|---|
| 📊 Executive Summary | KPI tiles, ML results, Pipeline log |
| 🗃 Cleaned Data | Full processed dataset (up to 50K rows) |
| 📐 Statistics | Min, Max, Mean, Median, Std, Sum per column |
| 📊 Charts | Bar chart, Pie chart, Line trend — auto-embedded |
| ⚡ Power BI Ready | Step-by-step Power BI connection guide |

### 🤖 ML Prediction
- Auto-detects revenue/target column
- Trains Random Forest Regressor
- Returns R² score, MAE, next-period prediction, growth %
- Shows top feature importances

### ⚡ Power BI Integration
Download the Excel report → Open in Power BI Desktop → Select "🗃 Cleaned Data" sheet → Load → Build your dashboard. Done.

---

## 🛠 Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python, Flask |
| Data Processing | Pandas, NumPy |
| Machine Learning | Scikit-learn (Random Forest) |
| Excel Generation | OpenPyXL (charts + formatting) |
| Frontend | HTML, CSS, JavaScript, Chart.js |
| Deployment | Docker, Render.com |

---

## 📁 Project Structure

```
bi-reporter/
├── app.py           # Flask backend — ETL + ML + Excel generator
├── index.html       # Frontend dashboard UI
├── requirements.txt # Python dependencies
├── Dockerfile       # Python 3.11.9 container
└── README.md
```

---

## ⚙️ Local Setup

### Prerequisites
- Python 3.11+

```bash
# Clone
git clone https://github.com/Alert9820/bi-reporter.git
cd bi-reporter

# Install
pip install -r requirements.txt

# Run
python app.py

# Open
http://localhost:10000
```

### Docker

```bash
docker build -t bi-reporter .
docker run -p 10000:10000 bi-reporter
```

---

## 🌐 Deploy on Render

1. Push to GitHub
2. Render → New → Web Service → Connect repo
3. Select **Docker** as runtime
4. Deploy ✅

---

## 🔌 API Endpoints

| Method | Endpoint | Description |
|---|---|---|
| `GET` | `/` | Serve dashboard UI |
| `POST` | `/upload` | Upload CSV/Excel file |
| `GET` | `/results/{job_id}` | Get ETL + ML results |
| `GET` | `/download/{job_id}` | Download Excel report |
| `GET` | `/health` | Health check |

---

## 📊 Supported Columns (Auto-Detected)

| Pattern | Column Names |
|---|---|
| Revenue | `revenue`, `sales`, `income`, `amount`, `price` |
| Cost | `cost`, `expense`, `spend` |
| Category | `category`, `product`, `region`, `department` |
| Date | `date`, `month`, `period`, `year` |
| Quantity | `qty`, `units`, `quantity`, `volume` |

---

## 💡 Use Cases

- Data Analysts — Quick reporting from raw datasets
- Business Teams — Upload sales data, get formatted report instantly
- Power BI Users — Pre-process and structure data for dashboards
- Students — Learn automated reporting and ML integration

---

## 👨‍💻 Author

**Sunny Mukesh Chaurasiya**
Portfolio project demonstrating: ETL automation · Excel report generation · ML prediction · Power BI integration · Flask API · Docker deployment

---

## 📄 License

MIT License

---

<div align="center"><strong>📊 BI Reporter — Because reports should write themselves.</strong></div>

