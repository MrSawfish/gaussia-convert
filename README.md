# 🚀 Data Conversion of Cancer Growth progression from Excel to Access Database Format

The cancer lab I worked at tracks cancer by gaussia signal in the blood of mice. This data was scattered across excel sheets per different cancer lines; this program seeks to consolidate it into one database. 

## 🛠️ Features

- **Core Feature 1:**
- **Core Feature 2:** Brief description of its value.
- **Automation:** Briefly explain any continuous processes.

## 💻 Tech Stack

- **Code:** Python package using pandas and openpyxyl libraries
- **Database:**Microsoft Access

## 🚀 Getting Started

### Prerequisites

What you need to install before running the software:
- Python 3.13.12 (Openpyxl needs to work with it)
- pandas
- Openpyxl 3.1.5

### Installation

1. Clone the repository:
```bash
   git clone https://github.com/your-username/your-repo-name.git
   cd your-repo-name
```

2. Install the required packages:
```bash
   pip install -r requirements.txt
```

   Or install them individually:
```bash
   pip install pandas openpyxl==3.1.5
```

## 📖 Usage Examples

## 📖 Usage

1. Place all the Excel files you want to consolidate into a single folder (e.g. `data/raw_excel/`).

2. Run the script, pointing it at that folder:
```bash
   python convert_to_dataframe.py --input data/raw_excel/ --output combined_data.csv
```

3. The script will:
   - Scan the folder for `.xlsx` files
   - Parse each file's Gaussia signal data into a common dataframe format
   - Combine all files into a single consolidated dataset
   - Export the result (e.g. as `.csv`) ready for import into Microsoft Access

4. Import the resulting file into Access:
   - Open Microsoft Access → **External Data** → **New Data Source** → **From File** → **Text File**
   - Select the exported `.xlsx` and follow the import wizard to map it into your database table

## 🤝 Contributing

Contributions make the open-source community an amazing place to learn, inspire, and create.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

Distributed under the MIT License. See `LICENSE` for more information.

## 📬 Contact

Your Name - [@your_twitter](https://twitter.com) - email@example.com
Project Link: [https://github.com](https://github.com)
