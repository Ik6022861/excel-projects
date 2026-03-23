# Excel VBA Automated Yearly Report Generator

## 📌 Project Overview

The **Excel VBA Automated Yearly Report Generator** is a macro-driven solution designed to standardize, format, and consolidate data from multiple worksheets into a single **Yearly Report** sheet.

This project demonstrates automation of repetitive Excel tasks using **Visual Basic for Applications (VBA)** and a user-friendly **UserForm interface** to enhance usability.

The system automatically:

* Inserts standardized headers into each worksheet
* Applies consistent formatting
* Calculates totals
* Copies data from multiple sheets
* Merges all records into a centralized **YEARLY REPORT**

---

## 🎯 Objectives

* Automate repetitive Excel reporting tasks
* Ensure consistent formatting across worksheets
* Reduce manual errors
* Improve reporting efficiency
* Provide a simple user interface for non-technical users

---

## 🧰 Technologies Used

* Microsoft Excel
* VBA (Visual Basic for Applications)
* Excel UserForms
* Macro Automation

---

## 📂 Workbook Structure

The workbook contains multiple worksheets representing different divisions:

* EAST RECORDS
* WEST RECORDS
* NORTH RECORDS
* SOUTH RECORDS
* YEARLY REPORT (Final consolidated report)

---

## 🖥 User Interface (UserForm)

A custom **UserForm** was created to allow users to interact with the system easily without running macros manually.

### Features:

* Dropdown to select a worksheet
* Button to add worksheet data
* Button to generate the final report

### UserForm Controls:

* ComboBox — Select worksheet
* Button — Add Worksheet
* Button — Run Report

---

## ⚙️ Macros Description

### 1. InsertHeaders()

This macro:

* Inserts a new header row
* Adds column titles

Headers added:

Division
Category
Jan
Feb
Mar
Total

---

### 2. FormatHeaders()

This macro:

* Applies formatting to headers
* Sets font style and size
* Applies background color
* Formats numeric values as Currency
* Adjusts column widths

---

### 3. AutomateTotalSum()

This macro:

* Finds the last value in the **Total** column
* Inserts a SUM formula below the data
* Automatically calculates the total amount

Formula used:

=SUM(F2:LastCell)

---

### 4. LoopYealryReport()

This is the main automation macro.

It:

1. Loops through all worksheets
2. Skips the **YEARLY REPORT** sheet
3. Inserts headers
4. Formats data
5. Calculates totals
6. Copies data
7. Pastes data into the YEARLY REPORT sheet
8. Generates the final consolidated report

---

## 🔄 Process Flow

User clicks:

RunReport Button
↓
Loop through worksheets
↓
Insert Headers
↓
Format Data
↓
Calculate Totals
↓
Copy Data
↓
Paste into YEARLY REPORT
↓
Generate Final Report

---

## 🚀 How to Run the Project

Step 1: Open the Excel workbook

Step 2: Enable Macros

Step 3: Open the UserForm

Press:

RunReport

The system will automatically:

* Process all worksheets
* Format data
* Calculate totals
* Generate the Yearly Report

---

## 📊 Example Output

The final **YEARLY REPORT** sheet will contain:

Division | Category | Jan | Feb | Mar | Total

All records from every worksheet will be combined into one structured report.

---

## 💡 Key Skills Demonstrated

* VBA Automation
* Excel Macro Development
* Looping Through Worksheets
* Dynamic Range Handling
* UserForm Development
* Data Consolidation
* Report Automation
* Business Process Automation

---

## 🏁 Conclusion

This project showcases how Excel VBA can be used to automate data processing and reporting tasks efficiently. The use of a UserForm improves usability and makes the solution suitable for business users with minimal technical knowledge.

The automation reduces manual work, improves consistency, and speeds up report generation.

---
