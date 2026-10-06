# 📊 ZIPTO Dark Store Down – Noida Cluster Review

### Data Analytics | Excel | Power Query | Power Pivot | DAX | VBA

---

## 📌 Project Overview

This project is a **Data Analytics and Business Intelligence case study** focused on analysing the performance of **Zipto's dark stores in the Noida cluster**.

The objective was to identify the stores facing financial and operational problems, understand the factors associated with poor performance, and provide data-driven recommendations for business improvement.

The project was developed using:

- Microsoft Excel
- Power Query
- Power Pivot
- DAX
- PivotTables
- PivotCharts
- VBA

The analysis is divided into three major areas:

1. **Store Performance & Profitability**
2. **Delivery & Fulfilment**
3. **Products & Promotions**

---

# 🎯 Business Problem

The main business question addressed in this project was:

> **Which stores should we fix first, what exactly is going wrong, and what can be improved?**

To answer this question, the analysis follows the approach:

**Observation → Investigation → Insight → Business Action**

The project focuses on understanding:

- Which stores are profitable and which are loss-making?
- Where are return rates unusually high?
- Which locations have poor delivery performance?
- What products and categories are generating revenue?
- How are promotions affecting basket value?
- Which stores should be prioritised for further investigation?

---

# 📥 Interactive Excel Dashboard

The complete Excel workbook contains:

- All 3 analytical dashboards
- Power Query transformations
- Power Pivot data model
- DAX measures
- PivotTables
- PivotCharts
- VBA dashboard navigation

### 👉 [View / Download the Complete Excel Dashboard](https://drive.google.com/file/d/1Ma2h1YPDeGhTJQPc3GEJBvrrooyHWjKJ/view?usp=drive_link)

> **Note:** The `.xlsm` workbook is hosted on Google Drive because its size is approximately 41 MB.

---

# 📊 Dashboard 1 – Store Performance & Profitability

## Business Question

> **Which stores are making money and which stores are losing money?**

The first dashboard focuses on store-level financial performance.

### Key Metrics

- Total Orders
- Total Contribution
- Store P&L
- Two-Month Store Rent
- Return Rate
- Contribution Loss from Returns

### Key Findings

A total of **52,133 unique orders** were analysed.

The analysis showed:

- Total Contribution: approximately **₹30.26 lakh**
- Two-Month Store Rent: approximately **₹34.76 lakh**
- Cluster P&L: approximately **-₹4.50 lakh**
- Loss-making Stores: **4**

The four loss-making stores were:

- S03
- S05
- S07
- S09

---

## 🔴 Priority Store 1 – S07

### Sector 44 Metro

S07 was identified as the most critical store.

Key indicators:

- Two-month P&L: approximately **-₹6.59 lakh**
- Return Rate: **14.01%**
- Contribution Loss from Returns: approximately **₹3.74 lakh**
- On-Time Delivery Rate: **27.8%**

The combination of poor profitability, high returns and weak delivery performance makes S07 the highest-priority store for investigation.

---

## 🟠 Priority Store 2 – S03

### Sector 137 Riverside

S03 was identified as the second priority store.

Key indicators:

- Two-month P&L: approximately **-₹3.12 lakh**
- Return Rate: **7.94%**
- Contribution Loss from Returns: approximately **₹1.99 lakh**
- On-Time Delivery Rate: **25.8%**

---

## 💡 Dashboard 1 Insight

S07 and S03 stand out because they combine:

**Negative profitability + High returns + Weak delivery performance**

However, the analysis does **not** assume that returns or late deliveries directly caused the losses.

These indicators are treated as signals requiring deeper investigation.

---

# 🚚 Dashboard 2 – Delivery & Fulfilment

## Business Question

> **Are we delivering orders on time?**

The second dashboard focuses on delivery performance across the Noida cluster.

### Key Metrics

- Total Trips
- On-Time Delivery Rate
- Late Delivery Rate
- Cancelled Trips
- Average Delivery Time
- Delivery Status Distribution

---

## 📌 Overall Delivery Performance

A total of **52,445 trip records** were analysed.

Overall performance:

| Delivery Status | Percentage |
|---|---:|
| On Time | 37.3% |
| Late | 60.7% |
| Cancelled | 2.0% |

The results show that late deliveries are a significant operational issue across the cluster.

---

## 📍 Lowest On-Time Locations

| Location | On-Time Rate |
|---|---:|
| Sector 137 Riverside | 25.8% |
| Sector 44 Metro | 27.8% |

These two locations correspond to:

- **S03 – Sector 137 Riverside**
- **S07 – Sector 44 Metro**

---

## 💡 Dashboard 2 Insight

An important pattern appears when we compare Dashboard 1 and Dashboard 2:

The stores with the largest profitability problems are also among the locations with the weakest delivery performance.

This creates an important **investigation hypothesis**:

> Poor delivery performance may be associated with weaker store economics.

However, this analysis does **not claim causation**.

To validate the relationship, further analysis should be performed by:

- Delivery Partner
- Route
- Time Period
- Return Reason
- Store
- Delivery Status

---

# 🛒 Dashboard 3 – Products & Promotions

## Business Question

> **What are we selling, and what are the discounts costing us?**

The third dashboard focuses on product performance, category revenue, profitability and promotions.

### Key Metrics

- Delivered Gross Revenue
- Total Units Sold
- Promo Orders
- Average Discount per Promo Order
- Category Revenue
- Gross Profit
- Gross Profit Margin
- Product Sales

---

## 📌 Key Product & Promotion Metrics

The analysis showed:

- Delivered Gross Revenue: approximately **₹2.67 crore**
- Valid Units Sold: approximately **2.60 lakh**
- Promo Orders: approximately **12,066**
- Average Discount per Delivered Promo Order: approximately **₹129.64**

---

# 🏷️ Category Performance

Revenue was analysed at the **item level** because one order can contain multiple products and categories.

Major revenue-generating categories included:

- Beverages
- Snacks & Packaged
- Fruits & Vegetables
- Personal Care
- Dairy & Eggs
- Household
- Meat & Seafood
- Bakery

### Important Analytical Point

The category analysis uses **item-level revenue**, while the overall delivered gross revenue KPI uses **order-level gross amount**.

Therefore, these figures should not be expected to match exactly.

Using item-level data prevents incorrect duplication when an order contains multiple products.

---

# 🏆 Top Products

The dashboard also analyses the **Top 10 Products by Units Sold**.

The top 10 products had relatively similar sales volumes.

The range was approximately:

**5,961 – 6,089 units**

This indicates that there was no single extremely dominant product among the top 10 based on unit volume.

---

# 🎟️ Promotion Analysis

The promotion analysis compares the basket value of promotional and non-promotional orders.

Promo orders had a higher average gross basket value.

However, after discounts were applied, the net basket value became much closer to the non-promo basket value.

### 💡 Business Insight

Promotions should not be evaluated only on gross sales.

A better approach is to evaluate:

**Gross Basket → Discount → Net Basket → Contribution**

This helps determine whether a promotion is actually creating profitable incremental value.

---

# 🔍 Data Cleaning & Transformation

The project involved multiple datasets that required cleaning and transformation before analysis.

### Power Query was used for:

- Removing duplicates
- Handling missing values
- Correcting data types
- Cleaning text fields
- Transforming raw datasets
- Creating calculated columns
- Preparing analysis-ready tables

---

# 🧮 Product-Level Calculations

The following calculations were created for product-level analysis.

### Item Revenue

```text
Item Revenue = Quantity × Unit Price
