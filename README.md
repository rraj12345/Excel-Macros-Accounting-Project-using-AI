# 📄 Invoice PDF Generator & Factory Data Extraction (Excel VBA)

## 📌 Project Overview

This project automates **invoice PDF generation** and **factory data extraction** using **Excel VBA**.
It is designed to reduce manual work in diamond/factory operations by automatically creating structured folders, generating invoice PDFs, and extracting factory-related sections from raw data.

The system reads data from structured Excel sheets and performs automation with a single **Generate button**.

---

# 🎯 Business Problem

In factory and diamond operations, teams often face:

* Manual invoice PDF creation
* Unstructured data stored in Excel sheets
* Time-consuming factory pending data extraction
* Risk of human error in folder creation and file naming
* Difficulty in organizing invoices party-wise and date-wise

This project solves these problems through **VBA automation**.

---

# 🚀 Features

### 1️⃣ Invoice PDF Generator

* Reads **Party Name, Invoice Number, and Date**
* Automatically creates folder structure:

```
D:\PARTY
   ├── Party Name
   │      ├── Date
   │           ├── Date_InvoiceNo.pdf
```

* Exports **INVOICE sheet as PDF**
* Opens PDF automatically after generation
* Shows success message

---

### 2️⃣ Factory Data Extraction

Automatically extracts:

* Factory Pending Detail
* Factory Repairing Pending Detail

From **DATA DUMP sheet** and creates:

* Factory Pending Sheet
* Factory Repairing Sheet

This helps in quick factory analysis and reporting.

---

# 🧾 Excel Sheets Structure

### 📄 INVOICE Sheet

Contains:

* Party Name
* Address
* GST Details
* Invoice Number
* Date
* Diamond Details
* Total Summary
* Generate Button

Used to create PDF invoice.

---

### 📊 DATA DUMP Sheet

Contains raw operational data:

* Polish details
* Process pending
* Factory pending detail
* Factory repairing detail
* Totals

Used for data extraction.

---

# 🛠️ VBA Macros

## 🔹 CreateInvoicePDF()

### Function

* Validates required fields
* Creates folders
* Generates PDF
* Saves invoice

### Required Fields

```
C4 → Party Name
I4 → Invoice Number
I5 → Date
```

### Output

```
D:\PARTY\PartyName\Date\Date_InvoiceNo.pdf
```

---

## 🔹 ExtractFactorySections()

### Function

Finds and extracts:

```
FACTORY PENDING DETAIL → Factory Pending Sheet
FACTORY REPAIRING PENDING DETAIL → Factory Repairing Sheet
```

### Process

* Finds section in DATA DUMP
* Copies data till TOTAL
* Creates new sheets
* Pastes data automatically

---

# 🖥️ Workflow

### Step 1

Update data in **DATA DUMP**

↓

### Step 2

Invoice sheet gets updated

↓

### Step 3

Click **GENERATE button**

↓

### Step 4

PDF created in:

```
D:\PARTY
```

↓

### Step 5

Factory data extracted automatically

---

# 📂 Project Structure

```
Excel VBA Invoice Automation
│
├── DATA DUMP Sheet
├── INVOICE Sheet
├── Factory Pending Sheet
├── Factory Repairing Sheet
│
├── CreateInvoicePDF VBA
├── ExtractFactorySections VBA
│
└── README.md
```

---

# 🧰 Tools & Technologies

* Microsoft Excel
* VBA (Visual Basic for Applications)
* File System Automation
* PDF Export
* Data Extraction Logic

---

# 💡 Key VBA Concepts Used

* Worksheet referencing
* Range validation
* Folder creation using `MkDir`
* File path handling
* ExportAsFixedFormat (PDF)
* Find function
* Dynamic range extraction
* Error handling
* Sheet creation & clearing

---

# 📸 Screens

## Invoice Sheet

* Delivery challan format
* Generate button
* Diamond details
* GST and party details

## Data Dump Sheet

* Polish summary
* Process pending
* Factory pending
* Factory repairing pending

---

# ⚙️ How to Use

### 1️⃣ Download Excel File

Clone repository

```
git clone https://github.com/yourusername/invoice-vba-automation.git
```

---

### 2️⃣ Open Excel

Enable:

```
Enable Editing
Enable Macros
```

---

### 3️⃣ Fill Invoice Details

```
Party Name
Invoice Number
Date
Diamond Details
```

---

### 4️⃣ Click Generate

PDF will be created automatically.

---

# 📈 Benefits

✔ Saves time
✔ Eliminates manual PDF creation
✔ Organized folder structure
✔ Automated factory reporting
✔ Easy to use
✔ Scalable for multiple invoices
✔ Reduces human error

---

# 🔮 Future Improvements

* Auto email PDF to party
* Auto invoice numbering
* Power BI integration
* Dashboard for factory performance
* Database connection
* Multi-invoice batch generation

If you want, I can **also create a professional GitHub repo structure (with badges, license, screenshots section, and animated GIF of Generate button)**.
