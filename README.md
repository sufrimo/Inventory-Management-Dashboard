# Inventory-Management-Dashboard
An Excel-based inventory management dashboard that tracks expired products, stockouts, supplier lead times, fulfillment times, and warehouse utilization. Utilizing Power Query and Power Pivot, it enables powerful data processing and analysis combined with interactive visualizations for real-time insights. The system also automates purchase order generation, supports printing, and converts orders to PDF for easy sharing and record-keeping.

## 🚀 Project Goal

The objective of this project is to develop a scalable and comprehensive inventory management dashboard that enables stakeholders to:

- Monitor expired products  
- Identify the most frequently stocked-out products  
- Analyze average supplier lead times  
- Track average order fulfillment time  
- Evaluate warehouse space utilization
- Streamline procurement through automated purchase order generation, printing, and PDF conversion using VBA

---

## 🧩 Features

- Power Query for data preparation  
- Conditional logic to identify expired and frequently stocked-out items  
- Custom columns for performance tracking  
- PivotTables and PivotCharts for dynamic insights  
- Interactive dashboard with slicers and charts  
- Power Pivot data model integration
- VBA automation for purchase order creation, printing, and PDF export to facilitate procurement processes

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
  
### 5. VBA Script and Form Control
- Write VBA macros to automate key tasks:

  - **CreatePurchaseOrder** — Generates purchase orders automatically.

  - **PrintAllPurchaseOrders** — Prints all purchase orders directly from Excel.

  - **ConvertPurchaseOrdersToPDF** — Exports purchase orders as PDF files for easy sharing and archiving.
- Insert Form Controls (buttons) on the dashboard and assign these macros:

  - A **“Create Purchase Order”** Form Control to run `CreatePurchaseOrder`.

  - A **“Print All”** Form Control to run `PrintAllPurchaseOrders`.

  - A **“Convert to PDF”** Form Control to run `ConvertPurchaseOrdersToPDF`.

---

## 📸 Sample Visuals

![Dashboard Snapshot](https://github.com/sufrimo/Inventory-Monitoring-with-Automated-Purchase-Order-Generation-Using-VBA/raw/main/Inventory%20Dashboard.jpg)


---

## 🧠 Tools & Technologies

- Microsoft Excel (Power Query + Power Pivot + VBA Macro)  
- PivotTables & Charts
-   

---

## ✅ To-Do / Future Enhancements

🔄 Integrate with Sales Records for Live Inventory
Link inventory levels with real-time or periodic sales data to track stock depletion accurately.

🚨 Alert When Stocks Fall to Reorder Level
Use conditional formatting or formulas to trigger low-stock alerts and prevent stockouts.

📦 Integrate Order Quantity Data for Automatic Purchase Orders
Pull order quantity and reorder level data from purchase requisitions, sales forecasts, or historical order records to automatically generate purchase orders when stock reaches reorder thresholds.

---

## 📃 License

This project is open-source under the MIT License.
