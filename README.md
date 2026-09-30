# Retail Sales Data Cleaning using Excel

## 📌 Project Overview

This project focuses on cleaning and preparing a retail sales dataset using Microsoft Excel.

The objective was to identify data quality and formatting issues in the raw retail sales dataset, apply appropriate Excel-based cleaning techniques, and prepare a structured dataset suitable for further analysis.

## 🎯 Project Objectives

- Clean and standardize the raw retail sales dataset
- Improve date and time-related fields
- Standardize phone number formatting
- Improve postal code data integrity
- Remove irrelevant or unnecessary columns
- Create useful derived fields
- Create unique customer identifiers
- Standardize financial fields
- Prepare the final dataset for analysis

## 🛠️ Tools Used

- Microsoft Excel
- Excel Formulas
- Cell Formatting
- Data Cleaning
- Data Validation

## 🔄 Data Cleaning Process

### 1. Date Formatting

The Order Date field was changed from a long date format to a short date format to improve readability and consistency.

### 2. Quarter Name

A QTR Name column was added to provide a clearer representation of the quarter.

#### Formula used:

`="Q"&Cell Reference`

### 3. Month Name

A Month Name column was created by converting the Month ID into the corresponding month name.

#### Formula used:

`=TEXT(DATE(2026,J2,1),"mmmm")`

### 4. Phone Number Cleaning

Phone numbers were cleaned by extracting numeric characters from the original phone-number field.

#### Formula used:

`=TEXTJOIN("",TRUE,IFERROR(MID(Q2,SEQUENCE(LEN(Q2)),1)*1,""))`

A separate cleaned phone-number column was then created to standardize valid 10-digit phone numbers.

#### Formula used:

`=IF(LEN(R2)=10,TEXT(R2,"(000)000-0000"),"NA")`

### 5. Address Cleaning

The Address Line 2 column was removed because it was mostly empty and considered non-relevant for the analysis.

### 6. Postal Code Formatting

The Postal Code column was changed from Number format to Text format because some postal codes contained alphanumeric values.

### 7. Currency Formatting

Currency formatting was applied to the Sales and Price columns to improve readability and consistency.

### 8. Removal of Irrelevant Columns

The MSRP column was removed because it was considered non-relevant to the project.

Non-relevant Phone Number columns, QTR_ID and Month ID were also removed from the final dataset.

### 9. Unique Customer ID

A unique Customer ID was created using Excel functions.

#### Formula used:

`=IF(COUNTIF($O$2:O2,O2)=1,MAX($P$1:P1)+1,VLOOKUP(O2,$O$1:P1,2,FALSE))`

### 10. Value-Only Conversion

After completing the cleaning and transformation process, the data was copied to a new sheet and pasted as values only.

This ensured that the final dataset contained the resulting values rather than dependent formulas.

### 11. Final Data Formatting

The final dataset was formatted as an Excel Table to improve structure, readability and usability.


## 📊 Before and After Cleaning

### Before Cleaning

The following screenshot shows the original retail sales dataset before the cleaning process.

![Before Cleaning](Data%20Before%20Cleaning.png)

### After Cleaning

The following screenshot shows the dataset after applying the cleaning and formatting steps.

![After Cleaning](Data%20after%20cleaning.png)

---

## 📁 Project Files

| File | Description |

| `retail-sales-data.csv` | Original retail sales dataset | 

| `retail-sales-cleaned-data.csv` | Cleaned and prepared retail sales dataset | 

| `Data Before Cleaning.png` | Dataset before the cleaning process | 

| `Data after cleaning.png` | Dataset after the cleaning process |

---

## 📈 Key Improvements

The cleaning process resulted in:

- More consistent date formatting
- Improved month and quarter readability
- Standardized phone number representation
- Improved postal code data integrity
- Removal of irrelevant fields
- Standardized currency presentation
- Creation of a unique customer identifier
- Removal of formula dependencies through value-only conversion
- A structured final dataset ready for analysis

## 💡 Skills Demonstrated

- Excel Data Cleaning
- Data Transformation
- Data Formatting
- Data Quality Improvement
- Excel Formula Application
- Data Standardization
- Data Validation
- Feature/Column Creation
- Data Preparation
- Analytical Thinking
