# 📂 Work Order Viewer

If you're using **QuickBooks** as your ERP system, this demo app shows you how to **track the status of each Work Order (WO)** in a centralized and automated way.

With this tool, you can easily:
- Check **whether a Work Order is completed**
  - If **not completed** → review the **inventory status** for required components
  - If **completed** → inspect the **output details and serial numbers**

---

## 📄 What You Need

To get started, you'll need the following files (all generated from QuickBooks or your production records):

1. `WO.pdf` — The Work Order file from QuickBooks  
2. `WO.doc` — The document listing serial numbers and production details  
3. `Inventory Stock Status.xlsx` — Inventory report exported from QuickBooks  
4. `Open Sales Order by Items.csv` — Sales order report by item from QuickBooks  

> ![Animation](https://github.com/user-attachments/assets/c10bff0c-027e-431a-9890-951c453dfaef)  
> *(Demo animation of the app in action)*

---

## 🔧 Features

- 🕵️ **Automatic File Monitoring**  
  Watches your folders for new or updated files using `watchdog`.

- 📄 **Work Order Extraction**  
  Parses part numbers and quantities from both **Word** and **PDF** documents.

- 🛢 **Persistent Storage with PostgreSQL**  
  Logs all work order and inventory information to a database.

- 📊 **Interactive Web Dashboard**  
  View and explore:
  - PDF versions of Work Orders  
  - Extracted Word document details (e.g., serial numbers)  
  - Inventory and component availability

- 🧩 **REST API Endpoint**  
  - `GET /api/word-files` → Returns up-to-date data from the Word files and their status.
