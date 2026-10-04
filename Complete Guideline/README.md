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

# 7.5 Validate Dates

Select:

```text
B2:B16
```

Go to:

```text
Data
→ Data Validation
```

Choose:

```text
Allow: Date
```

Set the appropriate date range.

The provided dates are:

```text
01-09-2026
02-09-2026
03-09-2026
04-09-2026
05-09-2026
```

### Result

All dates are valid.

### Answer

```text
All Date values are valid.
```

---

# 7.6 Validate Food Categories

The categories used are:

```text
Food
Fast Food
Beverage
```

Select:

```text
D2:D16
```

Then:

```text
Data
→ Data Validation
→ Allow: List
```

Enter:

```text
Food,Fast Food,Beverage
```

Click **OK**.

### Result

All categories are valid.

---

# 7.7 Validate Customer Types

Valid customer types are:

```text
Student
Teacher
Staff
```

Select:

```text
G2:G16
```

Then:

```text
Data
→ Data Validation
→ Allow: List
```

Enter:

```text
Student,Teacher,Staff
```

### Result

All customer types are valid.

---

# 7.8 Validate Payment Methods

Valid payment methods are:

```text
Cash
bKash
Nagad
```

Select:

```text
H2:H16
```

Then:

```text
Data
→ Data Validation
→ Allow: List
```

Enter:

```text
Cash,bKash,Nagad
```

### Result

All payment methods are valid.

---

# 7.9 Overall Validation Result

| Validation Check          | Result |
| ------------------------- | -----: |
| Missing Values            |      0 |
| Duplicate Transaction IDs |      0 |
| Invalid Quantity          |      0 |
| Invalid Unit Price        |      0 |
| Invalid Dates             |      0 |
| Invalid Categories        |      0 |
| Invalid Customer Types    |      0 |
| Invalid Payment Methods   |      0 |

### Final Validation Status

```text
The provided dataset passed all basic validation checks.
```

---

# 8. Stage 4 — Data Storage

The processed dataset should be stored in an Excel workbook.

## Recommended File Name

```text
University_Canteen_Sales_Data_Processing.xlsx
```

## Recommended Sheet Structure

### Sheet 1: Raw Data

Contains the original collected data.

```text
Raw Data
```

### Sheet 2: Processed Data

Contains:

* Cleaned data
* Total Sales
* Calculated values

```text
Processed Data
```

### Sheet 3: Analysis

Contains:

* Total revenue
* Average transaction
* Item-wise analysis
* Category-wise analysis
* Customer-wise analysis
* Payment-wise analysis
* Daily sales

```text
Analysis
```

### Sheet 4: Charts

Contains the final visualizations.

```text
Charts
```

---

# 9. Stage 5 — Data Processing

Now we calculate **Total Sales** for each transaction.

## Formula

```text
Total Sales = Quantity Sold × Unit Price
```

In Excel:

```text
I2 = E2*F2
```

Press **Enter**.

Then drag the formula from:

```text
I2
```

down to:

```text
I16
```

---

## Calculated Total Sales

| Transaction | Quantity | Unit Price | Total Sales |
| ----------- | -------: | ---------: | ----------: |
| T001        |       80 |        120 |       9,600 |
| T002        |       50 |        100 |       5,000 |
| T003        |       70 |         60 |       4,200 |
| T004        |       35 |        180 |       6,300 |
| T005        |       65 |        120 |       7,800 |
| T006        |       45 |         80 |       3,600 |
| T007        |       90 |         60 |       5,400 |
| T008        |       55 |        110 |       6,050 |
| T009        |       60 |        100 |       6,000 |
| T010        |       40 |        180 |       7,200 |
| T011        |       75 |        120 |       9,000 |
| T012        |       85 |         40 |       3,400 |
| T013        |       50 |        110 |       5,500 |
| T014        |       55 |        100 |       5,500 |
| T015        |      100 |         60 |       6,000 |

---
# 10. Stage 6 — Data Analysis

Now we analyze the processed sales data.

---

## 10.1 Total Revenue

### Formula

```excel
=SUM(I2:I16)
```

### Answer

```text
Total Revenue = 90,550
```

So, the total revenue generated by the 15 transactions is:

**90,550**

---

# 10.2 Average Transaction Sales

### Formula

```excel
=AVERAGE(I2:I16)
```

### Answer

```text
Average Transaction Sales = 6,036.67
```

Therefore, the average revenue per transaction is approximately:

**6,036.67**

---

# 10.3 Highest Transaction

### Formula

```excel
=MAX(I2:I16)
```

### Answer

```text
Highest Transaction = 9,600
```

Transaction:

```text
T001
```

Food Item:

```text
Chicken Biryani
```

---

# 10.4 Lowest Transaction

### Formula

```excel
=MIN(I2:I16)
```

### Answer

```text
Lowest Transaction = 3,400
```

Transaction:

```text
T012
```

Food Item:

```text
Soft Drink
```

---

# 10.5 Food Item-wise Total Sales

Use:

```excel
=SUMIF(C2:C16,"Chicken Biryani",I2:I16)
```

Similarly, replace the food item name for other items.

### Result

| Food Item       | Total Revenue |
| --------------- | ------------: |
| Chicken Biryani |        26,400 |
| Burger          |        16,500 |
| Coffee          |        15,600 |
| Pizza           |        13,500 |
| Fried Rice      |        11,550 |
| Sandwich        |         3,600 |
| Soft Drink      |         3,400 |

### Answer

**Chicken Biryani** generated the highest total revenue:

```text
26,400
```

---
# 10.6 Food Item-wise Total Quantity Sold

Use:

```excel
=SUMIF(C2:C16,"Chicken Biryani",E2:E16)
```

### Result

| Food Item       | Total Quantity Sold |
| --------------- | ------------------: |
| Coffee          |                 260 |
| Chicken Biryani |                 220 |
| Burger          |                 165 |
| Fried Rice      |                 105 |
| Soft Drink      |                  85 |
| Pizza           |                  75 |
| Sandwich        |                  45 |

### Answer

The most sold food item by quantity is:

```text
Coffee
```

Total quantity sold:

```text
260
```

---

# 10.7 Category-wise Revenue

Use:

```excel
=SUMIF(D2:D16,"Food",I2:I16)
```

For Fast Food:

```excel
=SUMIF(D2:D16,"Fast Food",I2:I16)
```

For Beverage:

```excel
=SUMIF(D2:D16,"Beverage",I2:I16)
```

### Result

| Category  | Revenue |
| --------- | ------: |
| Food      |  37,950 |
| Fast Food |  33,600 |
| Beverage  |  19,000 |

### Answer

The highest revenue-generating category is:

```text
Food
```

Revenue:

```text
37,950
```

---

# 10.8 Category-wise Average Sales

For Food:

```excel
=AVERAGEIF(D2:D16,"Food",I2:I16)
```

For Fast Food:

```excel
=AVERAGEIF(D2:D16,"Fast Food",I2:I16)
```

For Beverage:

```excel
=AVERAGEIF(D2:D16,"Beverage",I2:I16)
```

### Result

| Category  | Average Sales |
| --------- | ------------: |
| Food      |         7,590 |
| Fast Food |         5,600 |
| Beverage  |         4,750 |

### Answer

The **Food** category has the highest average transaction sales:

```text
7,590
```

---

# 10.9 Customer Type-wise Sales

Use:

```excel
=SUMIF(G2:G16,"Student",I2:I16)
```

For Teacher:

```excel
=SUMIF(G2:G16,"Teacher",I2:I16)
```

For Staff:

```excel
=SUMIF(G2:G16,"Staff",I2:I16)
```

### Result

| Customer Type | Total Sales |
| ------------- | ----------: |
| Student       |      62,250 |
| Staff         |      18,100 |
| Teacher       |      10,200 |

### Answer

The highest sales came from:

```text
Students
```

Total:

```text
62,250
```

---