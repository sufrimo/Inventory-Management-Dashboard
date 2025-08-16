# Inventory-Management-Dashboard
An Excel-based inventory management dashboard that tracks expired products, stockouts, supplier lead times, fulfillment time, and warehouse utilization using Power Query, Power Pivot, and interactive visualizations.

## 🚀 Project Goal

The objective of this project is to develop a scalable and comprehensive inventory management dashboard that enables stakeholders to:

- Monitor expired products  
- Identify the most frequently stocked-out products  
- Analyze average supplier lead times  
- Track average order fulfillment time  
- Evaluate warehouse space utilization  

---

## 🧩 Features

- Power Query for data preparation  
- Conditional logic to identify expired and frequently stocked-out items  
- Custom columns for performance tracking  
- PivotTables and PivotCharts for dynamic insights  
- Interactive dashboard with slicers and charts  
- Power Pivot data model integration  

---

## 📊 Dashboard KPIs

| Metric                        | Description                                                |
|------------------------------|------------------------------------------------------------|
| 🕓 Average Supplier Lead Time | Time between order placement and delivery by supplier      |
| 📦 Most Frequent Stockout Product | Items with highest frequency of stockout incidents    |
| ⏰ Average Fulfillment Time   | Time from order placed to received                         |
| 🏚 Warehouse Utilization      | Space usage relative to capacity                           |
| ⛔ Expired Products           | Products that have passed their expiry date               |

---

## ⚙️ Project Workflow

### 1. Load Data via Power Query
- Connect Excel/CSV files using Power Query  
- Use “Load to Power Pivot” to create a data model  

### 2. Data Preparation
- Clean and transform raw data  
- Add **Custom Columns** and **Conditional Columns**:
  - Check if product expiry date `<` today  
  - Check if stock level is below reordering point  

### 3. Build PivotTables & PivotCharts
- Track KPIs using PivotTables  
- Use charts for visuals:
  - Create reorder list of each store  
  - Supplier lead time and fulfillment time  
  - Warehouse utilization  
  - Most stockout product of each store  
  - Flag expired products  

### 4. Create Interactive Dashboard
- Include slicers for **Store ID**  

---

## 📸 Sample Visuals

[![Dashboard Snapshot](./images/Screenshot%202025-08-16%20090356.jpg)](https://github.com/sufrimo/Inventory-Management-Dashboard/blob/main/Screenshot%202025-08-16%20090356.jpg
)

---

## 🧠 Tools & Technologies

- Microsoft Excel (Power Query + Power Pivot)  
- PivotTables & Charts  

---

## ✅ To-Do / Future Enhancements

- 🔄 **Integrate with Sales Records for Live Inventory**  
  Link inventory levels with real-time or periodic sales data to track stock depletion accurately.

- 🚨 **Alert When Stocks Fall to Reorder Level**  
  Use conditional formatting or formulas to trigger low-stock alerts and prevent stockouts.

- 📝 **Set Up Automatic Purchase Order List via VBA Macro**  
  Generate purchase order suggestions automatically based on reorder points and lead times, using Excel VBA.

---

## 📃 License

This project is open-source under the MIT License.
