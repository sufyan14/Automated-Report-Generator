# 📊 Flexible Excel Report Generator with Python

This Python script generates a **dynamic, professional-looking Excel report** from any CSV dataset. It allows grouping by one or more columns, summarizes numerical values, creates a bar chart, formats the worksheet, and saves the final report — all automated!

---

## 🔧 Features

- ✅ Group data by one or more columns (e.g., job title, location)
- ✅ Summarize numeric values (e.g., total salary)
- ✅ Sort and display top N results
- ✅ Export a styled Excel report with:
  - Formatted headers and cells
  - Bar chart visualization
  - Column auto-width
  - Total row
  - Custom title and subtitle


## 🐍 Requirements

- Python 3.7 or higher  
- Libraries:
  - `pandas`
  - `openpyxl`

Install dependencies using:

```bash
pip install pandas openpyxl
````

---

## 📁 Example Use Case

If your dataset contains `job_title`, `company_location`, and `salary`, this script will output an Excel file showing the **Top 10 job roles with the highest total salaries** grouped by job title and location. You can modify columns and output variable according to your dataset.

---


### Function Parameters:

| Parameter       | Description                                                  |
| --------------- | ------------------------------------------------------------ |
| `file_path`     | Path to your CSV dataset                                     |
| `group_by_cols` | List of column names to group by (e.g., job title, location) |
| `value_col`     | Column to summarize (e.g., salary)                           |
| `output_file`   | Name for the generated Excel file                            |
| `top_n`         | Number of top results to include                             |

---

## 📊 Output Highlights

* Top N rows sorted by the chosen numeric value
* A bar chart visualizing the results
* Automatically formatted and styled Excel sheet
* A total row at the bottom for quick insights

✅ **Sample Output File:**
![{D8198C9B-467F-4FF8-AA83-5CE3D9239995}](https://github.com/user-attachments/assets/9a5f06df-49fb-48c9-938d-472f1951c1bc)

