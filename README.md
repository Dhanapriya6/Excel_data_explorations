# Excel_data_explorations

# 📊 Sales Database Analysis in Excel
# Excel-Task
This repository contains my Microsoft Excel tasks and Practical projects focused on data analysis and visualization. It showcases my ability to clean, organize, analyze, and interpret data using Excel.


# 📊 Data Exploration Using Excel
## 📌 Project Overview

This project focuses on analyzing a product sales database using Microsoft Excel. The dataset contains information about products, brands, prices, quantities, categories, price ranges, days, country codes, and months.

The objective is to transform raw sales data into meaningful insights by performing data cleaning, exploratory data analysis, KPI calculations, and dashboard creation using Excel.

## 🎯 Project Objectives

- Analyze product prices and quantities.
- Identify top-performing products and brands.
- Compare product categories based on quantity and estimated value.
- Analyze product distribution across countries.
- Explore monthly product performance.
- Identify high-price and standard-price products.
- Detect duplicate records and improve data quality.
- Create an interactive Excel dashboard for data visualization.

## 🛠️ Tools & Technologies

- Microsoft Excel
- Excel Formulas and Functions
- Data Cleaning
- Pivot Tables
- Pivot Charts
- Conditional Formatting
- Slicers and Filters
- Interactive Dashboard

## 📂 Dataset Description

The dataset contains the following columns:

| Column Name | Description |
|---|---|
| Product ID | Unique identifier containing day, month, and country code |
| Product Name | Name of the product |
| Brand Name | Manufacturer or brand associated with the product |
| Price ($) | Price of the product in US dollars |
| Quantity | Recorded quantity for each product |
| Category | Product category, such as Electronics, Fashion, and Kitchen |
| Price Range | Classification of the product price |
| Day | Day associated with the record |
| Country Code | Country identifier, such as US, UK, IN, and CA |
| Month | Month associated with the record |

## 🔍 Project Workflow

### 1. Data Cleaning and Preparation

- Checked for duplicate records.
- Reviewed missing and inconsistent values.
- Standardized product names, brand names, and categories.
- Verified price and quantity data types.
- Checked day, month, and country code consistency.
- Reviewed the accuracy of price range classifications.

### 2. Exploratory Data Analysis (EDA)

- Analyzed product prices and quantities.
- Compared products across different brands.
- Examined category-wise product distribution.
- Analyzed monthly and country-wise records.
- Identified the highest-priced products.
- Investigated duplicate product records.

### 3. Sales and Product Analysis

- Calculated total recorded quantity.
- Calculated average product price.
- Estimated product value using Price × Quantity.
- Compared estimated value across brands and categories.
- Analyzed the distribution of products by price range.
- Examined product performance across countries and months.

### 4. Dashboard Development

- Created Pivot Tables to summarize the dataset.
- Used Pivot Charts to visualize product performance.
- Added KPI cards for important metrics.
- Included slicers for Month, Country Code, Category, and Brand Name.
- Designed an interactive dashboard for business reporting.

## 📈 Key Performance Indicators (KPIs)

- **Total Product Records:** Number of records in the dataset.
- **Total Quantity:** Sum of all recorded quantities.
- **Average Product Price:** Average price across all records.
- **Estimated Total Value:** Sum of Price × Quantity.
- **Top Product by Quantity:** Product with the highest total recorded quantity.
- **Top Brand by Estimated Value:** Brand with the highest estimated product value.
- **Top Category:** Category with the highest estimated product value.
- **Country-Wise Performance:** Comparison of records, quantities, and estimated values across countries.

## 🧮 Excel Formulas Used

### 1. Total Quantity

```excel
=SUM(E2:E1000)
```

### 2. Average Product Price

```excel
=AVERAGE(D2:D1000)
```

### 3. Total Product Records

```excel
=COUNTA(A2:A1000)
```

### 4. Estimated Product Value

Add a new column called `Estimated Value` and enter:

```excel
=D2*E2
```

Copy the formula down for all records.

### 5. Total Quantity by Product

If the product name is entered in cell L2:

```excel
=SUMIF(B:B,L2,E:E)
```

### 6. Total Estimated Value by Brand

If the brand name is entered in cell L2 and the Estimated Value column is K:

```excel
=SUMIF(C:C,L2,K:K)
```

*Adjust the cell references and ranges to match your actual Excel worksheet.*

## 💡 Key Business Questions

- Which product has the highest price?
- Which product has the highest total quantity?
- Which brand contributes the most estimated product value?
- Which category has the highest estimated value?
- Which country has the highest recorded quantity?
- Which month has the highest estimated product value?
- How are products distributed across price ranges?
- Are there duplicate records in the dataset?
- Which products appear in multiple countries?
- How do product prices vary across categories?

## 📊 Dashboard Components

The Excel dashboard can include:

- KPI cards for total quantity, average price, and estimated value.
- Bar chart showing estimated value by brand.
- Column chart comparing product categories.
- Pie or doughnut chart showing price range distribution.
- Monthly analysis chart.
- Country-wise quantity and estimated value chart.
- Top products comparison chart.
- Interactive slicers for Month, Country Code, Category, and Brand Name.

## 📌 Data Quality Considerations

- The dataset contains repeated records, including Laptop Bag by Samsonite and Laptop by HP. These should be checked before removing duplicates.
- Headphones by Bose appears with different prices in two records. Both records should be retained unless the source confirms an error.
- Product ID values contain day, month, and country code, so they may not be globally unique transaction identifiers.
- The dataset does not contain an explicit revenue or sales amount column.

## 📊 Key Insights

The analysis is designed to identify:

- Best-performing products and brands.
- Categories contributing the most estimated value.
- Country-wise product distribution.
- Monthly patterns in recorded quantities and estimated value.
- Differences between high-price and standard-price products.
- Data quality issues that may affect reporting.

*Add verified numerical findings and charts after completing the Excel analysis.*

## 📷 Dashboard Preview

Add a screenshot of your completed Excel dashboard below.

![Sales Database Analysis Dashboard](dashboard/sales_dashboard.png)

## 🚀 How to Use

1. Clone or download this repository.
2. Open the Excel workbook in Microsoft Excel.
3. Review and clean the dataset.
4. Explore the Pivot Tables and charts.
5. Open the dashboard to view the KPIs.
6. Use slicers and filters to analyze products, brands, categories, months, and countries.

## 📚 Skills Demonstrated

- Data Cleaning and Preparation
- Exploratory Data Analysis (EDA)
- Excel Formulas and Functions
- Pivot Table Analysis
- Data Visualization
- Dashboard Development
- KPI Reporting
- Business Intelligence
- Analytical Thinking
- Data-Driven Decision Making

## 🎓 Project Outcome

This project demonstrates the use of Microsoft Excel to analyze a product database, summarize quantities and estimated product values, compare brands and categories, and visualize business information through interactive dashboards.

It showcases practical data analysis skills and the ability to transform raw data into structured reports that support business decision-making.




## 👤 Author

**V DHANAPRIYA**
**Aspiring Data Analyst**

This project is part of my Data Analytics learning journey, focusing on Excel functions, data exploration, and practical analysis.

## ⭐ Acknowledgment

Created for educational purposes to develop practical Microsoft Excel skills and establish a foundation for more advanced data analytics projects.


⭐ If you find this project useful, consider giving the repository a star!
