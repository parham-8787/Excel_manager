# 📊 Excel Manager

> A simple Python library to manage Excel files easily using [openpyxl](https://openpyxl.readthedocs.io/).

---

## 🚀 Features
- 📑 Create and load Excel workbooks
- ➕ Add, ✏️ rename, ❌ remove sheets
- 📥 Insert and display data
- 🔄 Copy data between Excel files

---

## 📦 Installation

Clone the repo and install requirements:

```bash
git clone https://github.com/parham-8787/Excel_manager.git
cd Excel_manager
pip install -r requirements.txt
```

```python
from excel_manager import Excel

# Create a new workbook
ex = Excel("test_file", "Sheet1", create_workbook="no")

# Insert some data
ex.insert_data([["Name", "Age"], ["Ali", 25]])

# Show inserted data
ex.show_data(1, 2, 1, 2)

# Create and rename a sheet
ex.creat_sheet("NewSheet")
ex.rename_sheet("NewSheet", "RenamedSheet")

# Show all sheets
ex.show_sheets()
```
