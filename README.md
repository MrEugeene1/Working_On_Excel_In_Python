<div align="center">

# 📊 Reading Excel Files with Pandas

**A hands-on beginner's guide to loading real-world Excel data into Python, and handling the messy parts.**

![Python](https://img.shields.io/badge/Python-3.14-3776AB?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-read__excel-150458?logo=pandas&logoColor=white)
![Jupyter](https://img.shields.io/badge/Jupyter-Notebook-F37626?logo=jupyter&logoColor=white)
![Level](https://img.shields.io/badge/level-beginner-brightgreen)

</div>

---

## 📚 Table of Contents

1. [What you'll learn](#-what-youll-learn)
2. [Project structure](#-project-structure)
3. [Setup](#-setup)
4. [Lesson walkthrough](#-lesson-walkthrough)
   - [1. Reading a simple Excel file](#1-reading-a-simple-excel-file)
   - [2. Files with multiple sheets](#2-files-with-multiple-sheets)
   - [3. Comments in the top rows](#3-comments-in-the-top-rows)
   - [4. Assignment: files spread across many folders](#4-assignment-files-spread-across-many-folders)
5. [Cheat sheet](#-cheat-sheet)
6. [Troubleshooting](#-troubleshooting)
7. [Practice ideas](#-practice-ideas)

---

## 🎯 What you'll learn

By the end of the notebook [`2. EXCEL FILES.ipynb`](2.%20EXCEL%20FILES.ipynb) you will be able to:

| ✅ Skill | 🔧 Tool |
|---|---|
| Read an `.xlsx` file into a DataFrame | `pd.read_excel()` |
| Pick a specific sheet from a workbook | `sheet_name=` |
| Stack several sheets into one table | `pd.concat()` |
| Skip junk/comment rows above the header | `skiprows=` |
| Find files across many sub-folders | `glob.glob()` |
| Save the combined result | `DataFrame.to_csv()` |

The data throughout is the **Big Mart Sales** dataset: product sales from retail outlets, with an `Outlet_Establishment_Year` column that we use to check which data was loaded.

---

## 🗂 Project structure

```text
reading_excel/
├── 2. EXCEL FILES.ipynb                       ← the lesson notebook
├── README.md                                  ← you are here
└── datasets/
    ├── big_mart_sales.xlsx                    ← clean, single-sheet file
    ├── big_mart_sales_with_multiple_sheets.xlsx  ← sheets: 1985, 1987, 1997
    ├── big_mart_sales_comments.xlsx           ← 3 comment rows above the header
    ├── merged_big_mart_sales.csv              ← output of the assignment
    └── multi-directory-excel/
        ├── 1985/1985.xlsx
        ├── 1987/1987.xlsx
        ├── 1997/1997.xlsx
        ├── 1998/1998.xlsx
        ├── 1999/1999.xlsx
        ├── 2002/2002.xlsx
        ├── 2004/2004.xlsx
        ├── 2007/2007.xlsx
        └── 2009/2009.xlsx
```

---

## 🚀 Setup

### 1. Install the packages

pandas needs a helper library called **`openpyxl`** to read `.xlsx` files. Without it you'll see
`ImportError: Missing optional dependency 'openpyxl'`.

**Inside the notebook** (recommended, because it installs into the same Python the notebook is running):

```python
%pip install pandas openpyxl
```

**From a terminal:**

```bash
python -m pip install pandas openpyxl
```

> 💡 **Why `python -m pip` and not just `pip`?**
> `python -m pip` runs pip *through* a specific Python interpreter, so packages always land where you expect.
> It also avoids the `pip.exe` launcher entirely, which helps on locked-down Windows machines (see [Troubleshooting](#-troubleshooting)).

### 2. Open the notebook

Open `2. EXCEL FILES.ipynb` in VS Code, JupyterLab, or Jupyter Notebook, and **make sure the working directory is this folder** so the relative paths like `datasets/big_mart_sales.xlsx` resolve.

---

## 🧭 Lesson walkthrough

### 1. Reading a simple Excel file

```python
import pandas as pd

data = pd.read_excel('datasets/big_mart_sales.xlsx')
data.head()          # peek at the first 5 rows
```

That's it. `read_excel` works just like `read_csv`, and returns a DataFrame.

---

### 2. Files with multiple sheets

Real workbooks often have **one sheet per year, region, or month**. Here's the catch:

> ⚠️ **By default `read_excel` reads only the FIRST sheet.** The rest are silently ignored.

```python
df = pd.read_excel('datasets/big_mart_sales_with_multiple_sheets.xlsx')
df.Outlet_Establishment_Year.unique()   # → only ONE year (the first sheet's) 😕
```

**Fix: name the sheet you want with `sheet_name`.**

```python
sheet_1985 = pd.read_excel('datasets/big_mart_sales_with_multiple_sheets.xlsx', sheet_name='1985')
sheet_1987 = pd.read_excel('datasets/big_mart_sales_with_multiple_sheets.xlsx', sheet_name='1987')
sheet_1997 = pd.read_excel('datasets/big_mart_sales_with_multiple_sheets.xlsx', sheet_name='1997')
```

**Then stack them into one table with `pd.concat`:**

```python
final_data = pd.concat([sheet_1985, sheet_1987, sheet_1997])
final_data.Outlet_Establishment_Year.unique()   # → 1985, 1987, 1997 ✅
```

> 🚀 **Pro tip:** `sheet_name=None` reads **every** sheet at once and returns a dictionary of `{sheet name: DataFrame}`:
>
> ```python
> all_sheets = pd.read_excel('datasets/big_mart_sales_with_multiple_sheets.xlsx', sheet_name=None)
> final_data = pd.concat(all_sheets.values(), ignore_index=True)
> ```
>
> `sheet_name` also accepts a position (`sheet_name=0` is the first sheet) or a list (`sheet_name=['1985', '1997']`).

---

### 3. Comments in the top rows

People love adding notes like *"Report generated by…"* above the actual table. pandas then treats that note as the header row, so you end up with columns named `Unnamed: 1`, `Unnamed: 2`, and so on, plus a few rows of `NaN`.

```python
messy = pd.read_excel('datasets/big_mart_sales_comments.xlsx')
messy.head()    # 😵 garbage header, NaN rows
```

**Fix: skip those rows with `skiprows`.**

```python
clean = pd.read_excel('datasets/big_mart_sales_comments.xlsx', skiprows=3)
clean.head()    # ✅ real column names
```

> 🔍 **How do I know how many rows to skip?** Open the file once without `skiprows`, look at `.head(10)`, and count the rows above the real header.
> Related options: `header=3` (says *which* row is the header) and `nrows=100` (read only the first N data rows).

---

### 4. Assignment: files spread across many folders

Sometimes data arrives as **one file per folder**, for example `multi-directory-excel/1985/1985.xlsx`, `.../1987/1987.xlsx`, and so on. The `glob` module finds paths that match a wildcard pattern.

**Step 1: list the folders**

```python
import glob

for directory in glob.glob('datasets/multi-directory-excel/*'):
    print(directory)
```

**Step 2: list the files inside each folder**

```python
for directory in glob.glob('datasets/multi-directory-excel/*'):
    for file in glob.glob(directory + '/*'):
        print(file)
```

**Step 3: read every file, combine, and save**

```python
data_frame_list = []

for directory in glob.glob('datasets/multi-directory-excel/*'):
    for file in glob.glob(directory + '/*'):
        data_frame_list.append(pd.read_excel(file))

final_data = pd.concat(data_frame_list)
final_data.to_csv('datasets/merged_big_mart_sales.csv', index=False)
```

> 🚀 **Shorter version:** `**` with `recursive=True` searches every sub-folder in one go:
>
> ```python
> files = glob.glob('datasets/multi-directory-excel/**/*.xlsx', recursive=True)
> final_data = pd.concat((pd.read_excel(f) for f in files), ignore_index=True)
> ```
>
> Using `*.xlsx` instead of `*` also stops temporary lock files (like `~$1985.xlsx`, created while a file is open in Excel) from sneaking in.

---

## 📝 Cheat sheet

| I want to… | Code |
|---|---|
| Read the first sheet | `pd.read_excel('file.xlsx')` |
| Read a sheet by name | `pd.read_excel('file.xlsx', sheet_name='1985')` |
| Read a sheet by position | `pd.read_excel('file.xlsx', sheet_name=0)` |
| Read **all** sheets (dict) | `pd.read_excel('file.xlsx', sheet_name=None)` |
| Skip the first N rows | `pd.read_excel('file.xlsx', skiprows=3)` |
| Say which row is the header | `pd.read_excel('file.xlsx', header=2)` |
| Read only some columns | `pd.read_excel('file.xlsx', usecols='A:D')` |
| Read only N rows | `pd.read_excel('file.xlsx', nrows=100)` |
| Combine DataFrames | `pd.concat([df1, df2], ignore_index=True)` |
| Find files in sub-folders | `glob.glob('folder/**/*.xlsx', recursive=True)` |
| Save as CSV (no index column) | `df.to_csv('out.csv', index=False)` |

---

## 🛠 Troubleshooting

### ❌ `'pip3.exe' was blocked by your organization's Device Guard policy`

You might see this on Windows even on a personal PC:

```text
'C:\Python314\Scripts\pip3.exe' was blocked by your organization's Device Guard policy.
Contact your support person for more info.
```

**What's going on:** `pip.exe` / `pip3.exe` are small *launcher* programs that pip generates when it's installed, and they aren't digitally signed. Windows security features such as **Smart App Control** (on Windows 11 Home) or **App Control / Device Guard** (on managed work PCs) block unsigned programs. `python.exe` itself *is* signed, so it's allowed to run.

**✅ Fix: don't use the launcher; run pip through Python:**

```powershell
python -m pip install pandas openpyxl
# or, if `python` isn't on your PATH:
C:\Python314\python.exe -m pip install pandas openpyxl
```

Inside Jupyter, `%pip install pandas openpyxl` works too.

**Alternative (not recommended):** on a personal Windows 11 PC you can turn Smart App Control off under
*Windows Security → App & browser control → Smart App Control settings*. This lowers your protection against malicious apps, and on some Windows versions you **can't turn it back on** without resetting Windows. On a work or school PC, ask your IT team instead; don't try to bypass the policy.

### ❌ `ImportError: Missing optional dependency 'openpyxl'`

Install it: `%pip install openpyxl`, then **restart the kernel**.

### ❌ `FileNotFoundError: datasets/big_mart_sales.xlsx`

The notebook isn't running from this folder. Check with:

```python
import os; print(os.getcwd())
```

### ❌ `SyntaxError` in the "list all the files" cell

In the notebook, the `f` of `for` has slipped onto the end of the comment line above it:

```python
# list all the files in the folderf
or directory in glob.glob(...):      # ← broken
```

Move the `f` back down so the line reads `for directory in glob.glob(...):`.

### ❌ Only one year shows up after reading a multi-sheet file

That's the default behaviour: only the first sheet is read. Use `sheet_name=` or `sheet_name=None` ([section 2](#2-files-with-multiple-sheets)).

---

## 🏋️ Practice ideas

Try these yourself to lock in what you learned:

1. 🔁 Rewrite the multi-sheet section using **`sheet_name=None`** and a single `pd.concat`.
2. 🏷 After concatenating, add a column recording which **file or sheet** each row came from. *(Hint: `df.assign(source=name)`.)*
3. 🔢 Use `ignore_index=True` in `pd.concat` and compare `final_data.index` before and after. Why does it matter?
4. 📏 Check that the merged data has as many rows as all the individual files added together.
5. 📈 Group the merged data by `Outlet_Establishment_Year` and calculate total sales per year.

---

<div align="center">

**Happy learning! 🐼✨**
*If this helped you, share it with someone else who is learning pandas.*

</div>
