# MIS Project: University Canteen Sales Data Processing

## Project Scenario

A university wants to develop a simple **Canteen Sales Information System**.

The university canteen collects sales information from different sources, but the raw data is not properly organized. Therefore, the data needs to be:

* Collected
* Entered into a spreadsheet
* Validated
* Stored
* Processed
* Analyzed
* Reported
* Visualized

This project demonstrates the complete **8-stage data processing workflow** using Microsoft Excel.

---

# Table of Contents

1. [Project Objectives](#1-project-objectives)
2. [8 Stages of Data Processing](#2-8-stages-of-data-processing)
3. [Raw Data](#3-raw-data)
4. [Data Fields](#4-data-fields)
5. [Stage 1 — Data Collection](#5-stage-1--data-collection)
6. [Stage 2 — Data Entry](#6-stage-2--data-entry)
7. [Stage 3 — Data Validation](#7-stage-3--data-validation)
8. [Stage 4 — Data Storage](#8-stage-4--data-storage)
9. [Stage 5 — Data Processing](#9-stage-5--data-processing)
10. [Stage 6 — Data Analysis](#10-stage-6--data-analysis)
11. [Stage 7 — Reporting](#11-stage-7--reporting)
12. [Stage 8 — Visualization](#12-stage-8--visualization)
13. [Final Answers](#13-final-answers)
14. [Important Excel Formulas](#14-important-excel-formulas)
15. [Final Excel Structure](#15-final-excel-structure)
16. [Final Checklist](#16-final-checklist)

---
# 1. Project Objectives

The main objectives of this project are:

* Organize raw canteen sales data.
* Validate the data for errors and inconsistencies.
* Calculate total sales for each transaction.
* Calculate total revenue.
* Calculate average transaction sales.
* Analyze sales by food item, category, customer type, payment method, and date.
* Identify the most sold and highest revenue-generating food items.
* Find the relationship between quantity sold and total sales.
* Create charts for better understanding of the sales data.

---

# 2. 8 Stages of Data Processing

The complete workflow is:

```text
Data Collection
       ↓
Data Entry
       ↓
Data Validation
       ↓
Data Storage
       ↓
Data Processing
       ↓
Data Analysis
       ↓
Data Reporting
       ↓
Data Visualization
```

---

# 3. Raw Data

The following information was collected from different sources:

```text
T001, 01-09-2026, Chicken Biryani, Food, 80, 120, Student, Cash
T002, 01-09-2026, Burger, Fast Food, 50, 100, Student, bKash
T003, 01-09-2026, Coffee, Beverage, 70, 60, Teacher, Cash
T004, 02-09-2026, Pizza, Fast Food, 35, 180, Student, Nagad
T005, 02-09-2026, Chicken Biryani, Food, 65, 120, Student, bKash
T006, 02-09-2026, Sandwich, Fast Food, 45, 80, Staff, Cash
T007, 03-09-2026, Coffee, Beverage, 90, 60, Student, bKash
T008, 03-09-2026, Fried Rice, Food, 55, 110, Student, Cash
T009, 03-09-2026, Burger, Fast Food, 60, 100, Teacher, Nagad
T010, 04-09-2026, Pizza, Fast Food, 40, 180, Student, bKash
T011, 04-09-2026, Chicken Biryani, Food, 75, 120, Staff, Cash
T012, 04-09-2026, Soft Drink, Beverage, 85, 40, Student, bKash
T013, 05-09-2026, Fried Rice, Food, 50, 110, Student, Nagad
T014, 05-09-2026, Burger, Fast Food, 55, 100, Staff, Cash
T015, 05-09-2026, Coffee, Beverage, 100, 60, Student, bKash
```

---

# 4. Data Fields

The raw data contains 8 fields.

| Column | Field          | Description                   |
| ------ | -------------- | ----------------------------- |
| A      | Transaction ID | Unique ID of each transaction |
| B      | Date           | Date of the transaction       |
| C      | Food Item      | Name of the food or beverage  |
| D      | Category       | Food category                 |
| E      | Quantity Sold  | Number of units sold          |
| F      | Unit Price     | Price of one unit             |
| G      | Customer Type  | Student, Teacher, or Staff    |
| H      | Payment Method | Cash, bKash, or Nagad         |

After processing, we will add:

| Column | Field       | Formula               |
| ------ | ----------- | --------------------- |
| I      | Total Sales | Quantity × Unit Price |

---

# 5. Stage 1 — Data Collection

## Possible Data Sources

At least 3 possible sources of canteen sales data are:

### 1. Canteen POS / Billing System

The canteen billing system can automatically store:

* Transaction ID
* Food item
* Quantity
* Unit price
* Total sales
* Payment method

### 2. Daily Sales Register

Canteen staff can maintain a daily sales register containing sales information.

### 3. Digital Payment Records

Payment records from:

* bKash
* Nagad
* Bank/payment systems

can provide information about digital transactions.

### Other Possible Sources

* Cash register
* Food ordering system
* Canteen inventory system
* Manual sales sheets

### Answer

**Three suitable sources are:**

1. Canteen POS/Billing System
2. Daily Sales Register
3. Digital Payment Records

---

# 6. Stage 2 — Data Entry

We will use **Microsoft Excel** to organize the raw data.

## Step 1: Open Excel

Open Microsoft Excel and create a new workbook.

## Step 2: Create Headers

Enter the following headers in Row 1:

```text
A1 = Transaction ID
B1 = Date
C1 = Food Item
D1 = Category
E1 = Quantity Sold
F1 = Unit Price
G1 = Customer Type
H1 = Payment Method
I1 = Total Sales
```

## Step 3: Enter the Data

Enter the 15 transactions from the raw data.

The data will occupy:

```text
A2:H16
```

## Step 4: Convert into an Excel Table

Select:

```text
A1:I16
```

Then:

```text
Insert → Table
```

or use:

```text
Ctrl + T
```

Select:

```text
My table has headers
```

Click **OK**.

This makes filtering, sorting, and analysis easier.

---
# 7. Stage 3 — Data Validation

Data validation checks whether the dataset contains incorrect, missing, duplicated, or invalid values.

We need to check:

* Missing data
* Duplicate Transaction IDs
* Quantity Sold
* Unit Price
* Dates
* Food Categories
* Customer Types
* Payment Methods

---

## 7.1 Check Missing Data

Select:

```text
A2:H16
```

Then:

```text
Home
→ Find & Select
→ Go To Special
→ Blanks
→ OK
```

### Result

There are **no missing values** in the provided dataset.

### Answer

```text
Missing Values = 0
```

---

# 7.2 Check Duplicate Transaction IDs

Transaction IDs should be unique.

Select:

```text
A2:A16
```

Then:

```text
Home
→ Conditional Formatting
→ Highlight Cells Rules
→ Duplicate Values
```

Select a formatting style and click **OK**.

### Result

No duplicate Transaction IDs were found.

### Answer

```text
Duplicate Transaction IDs = 0
```

---

# 7.3 Validate Quantity Sold

Quantity Sold should be a positive number.

Select:

```text
E2:E16
```

Then:

```text
Home
→ Conditional Formatting
→ Highlight Cells Rules
→ Less Than
```

Enter:

```text
1
```

### Result

All quantities are positive.

The minimum quantity is:

```text
35
```

The maximum quantity is:

```text
100
```

### Answer

```text
All Quantity Sold values are valid.
```

---

# 7.4 Validate Unit Price

Unit Price should also be greater than 0.

Select:

```text
F2:F16
```

Then:

```text
Home
→ Conditional Formatting
→ Highlight Cells Rules
→ Less Than
```

Enter:

```text
1
```

### Result

All unit prices are positive.

Prices range from:

```text
40 to 180
```

### Answer

```text
All Unit Price values are valid.
```

---