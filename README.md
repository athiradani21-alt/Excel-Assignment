# Excel Assignment 1 - Data Exploration

This repository contains my solution for my Assignment 1 -Data Exploration

---
## Objective
The Primary Aim of this Assignment is to transform the into a accurate data by cleaning,analysing.Moreover to extract insights from a raw data using arithametic and statiscal function like SUM,AVERAGE ,COUNT,COUNTIF, SUMIF,MIN, MAX,LEFT RIGHT AND MID

Dataset Overview
This Dataset Consist of 
* **Product ID**
* **Category** 
* **Price**




### 1) Sum, Count, Average
* **Total Price of all products:** Calculates the overall sum of the price column.
  * *Formula:* `=SUM(D2:D35)
* **Total Product Count:** Counts the total number of items/rows in the dataset.
  * *Formula:* `=COUNTA(D2:D35)
* **Average Price:** Computes the average price across all products.
  * *Formula:* `=AVERAGE(D2:D35)
  * * Use the fill function to fill the column

---

### 2) Min and Max
* **Minimum Price:** Finds the lowest-priced product in the dataset.
  * *Formula:* `=MIN(D2:D35)
  * * Use the fill function to fill the column.
* **Maximum Price:** Finds the highest-priced product in the dataset.
  * *Formula:* `=MAX(D2:D35)
  * * Use the fill function to fill the column.

### 3) IF Function (`Price Range` Column)
Categorizes products based on their price: products priced at **$500 or higher** are labeled as `'High Price'`, while everything else is marked as `'Standard Price'`.
* *Formula:* =IF(D2>=500,"High Price","Standard Price")
* * Use the fill function to fill the column

---

## 4) SUMIF and COUNTIF
* **Total Price for 'Electronics':** Sums the price of products belonging specifically to the 'Electronics' category.
  * *Formula:* =SUMIF(F2:F35,F2,D2:D35)
* **Count of Products under $100:** Counts how many products have a price lower than $100.
  * *Formula:* =COUNTIF(D2:D35,"<100")
  * * Use the fill function to fill the column.

---

## 5) Text Formatting (`LEFT`, `RIGHT`, `MID`)
Extracts specific sub-strings from the `'Product ID'` column into new columns:
* **Day (First 2 characters):** Extracts the first 2 characters from the Product ID to create the `Day` column.
  * *Formula:* =LEFT(A2,2)
  * Use the fill function to fill the column.
* **Country Code (Last 2 characters):** Extracts the last 2 characters from the Product ID to create the `Country Code` column.
  * *Formula:* =RIGHT(A2,2)
  * * Use the fill function to fill the column.
* **Month (4th to 6th characters):** Extracts 3 characters starting from the 4th character position of the Product ID to create the `Month` column.
  * *Formula:* `=MID(A2,4,3)
  * * Use the fill function to fill the column.
  
