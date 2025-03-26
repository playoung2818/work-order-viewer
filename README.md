# 📂 Work Order Viewer

This is a full-stack Python application that automates the tracking and display of work order files (PDF & Word) along with their real-time inventory status. It integrates:

- 📄 File monitoring
- 🧠 Document parsing (PDF & Word)
- 🗃️ PostgreSQL logging
- 🌐 Flask web & API server

---

## 🔧 Features

- **Automatic file monitoring** with `watchdog`
- **Extracts part numbers and quantities** from Word and PDF work orders
- **Stores data into PostgreSQL** for persistent logging
- **Serves an interactive dashboard** for:
  - Viewing work order PDFs
  - Displaying extracted Word file info
  - Showing real-time inventory status
- **REST API**: `GET /api/word-files` — returns current word file data and updates status

---

## 📁 Project Structure

