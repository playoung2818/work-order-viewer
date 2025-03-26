import os
import logging
from flask import Flask, render_template_string, request, send_file, abort
from flask_sqlalchemy import SQLAlchemy
import urllib.parse
import json
import requests # type: ignore
import pandas as pd
import re


# Flask Configuration
app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'postgresql://postgres:Czheyuan0227%40@localhost:5432/File_Log'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# Database Models
class PDFFileLog(db.Model):
    __tablename__ = 'pdf_file_log'
    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.String(255), nullable=False)
    file_name = db.Column(db.String(255), nullable=False)
    file_path = db.Column(db.String(255), nullable=False)

class WordFileLog(db.Model):
    __tablename__ = 'word_file_log'
    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.String(255), nullable=False)
    file_name = db.Column(db.String(255), nullable=False)
    product_details = db.Column(db.JSON, nullable=True)
    file_path = db.Column(db.String(255), nullable=False)


# 🏗 Load Inventory & Sales Order Data
sales_order_file = r"c:\Users\Admin\OneDrive - neousys-tech\Desktop\Open Sales Order\Open Sales Order 3_20_2025.CSV"
warehouse_inventory_file = r"c:\Users\Admin\OneDrive - neousys-tech\Desktop\QuickBook\WH01S_3_20.CSV"

try:
    df_sales_order = pd.read_csv(sales_order_file, encoding="ISO-8859-1")
except UnicodeDecodeError:
    df_sales_order = pd.read_csv(sales_order_file, encoding="latin1")

try:
    inventory_df = pd.read_csv(warehouse_inventory_file, encoding="ISO-8859-1")
except UnicodeDecodeError:
    inventory_df = pd.read_csv(warehouse_inventory_file, encoding="latin1")

# 📡 API URL
url = "http://192.168.60.121:5001/api/word-files"

# Fetch API data
response = requests.get(url)

if response.status_code == 200:
    api_data = response.json()
    
    # Extract "word_files" key if available
    if "word_files" in api_data and isinstance(api_data["word_files"], list):
        word_files_df = pd.DataFrame(api_data["word_files"])
    else:
        print("⚠️ Unexpected API format, creating an empty DataFrame.")
        word_files_df = pd.DataFrame(columns=["file_name", "order_id", "status"])
else:
    print(f"Error: {response.status_code} - {response.text}")
    word_files_df = pd.DataFrame(columns=["file_name", "order_id", "status"])

# # Debugging: Print API data structure
# print("Columns in word_files_df:", word_files_df.columns.tolist())
# print(word_files_df.head())

# Function to extract numeric WO Number (digits between first and second "-")
def extract_wo_number(order_id):
    match = re.search(r'-(\d+)-', order_id)  # Extract numeric part
    return match.group(1) if match else None

# Ensure "order_id" exists and apply extraction function
if "order_id" in word_files_df.columns:
    word_files_df["WO_Number"] = word_files_df["order_id"].apply(lambda x: extract_wo_number(x) if isinstance(x, str) else None)
else:
    print("⚠️ 'order_id' column missing, cannot extract WO_Number.")
    word_files_df["WO_Number"] = None  

# Convert WO_Number to string for merging
word_files_df["WO_Number"] = word_files_df["WO_Number"].astype(str)

# 🏗 Load Open Sales Order File
try:
    df_sales_order = pd.read_csv(sales_order_file, encoding="ISO-8859-1")
except UnicodeDecodeError:
    df_sales_order = pd.read_csv(sales_order_file, encoding="latin1")

# Rename columns for clarity
df_sales_order.rename(columns={'Unnamed: 0': 'Component', 'Num': 'WO_Number', 'Qty': 'Required_Qty'}, inplace=True)

# Fill down the Component column
df_sales_order['Component'] = df_sales_order['Component'].ffill()

# Standardize component names (strip spaces and lowercase for consistency)
df_sales_order["Component"] = df_sales_order["Component"].astype(str).str.strip().str.lower()

# Remove "Total" rows to avoid duplicate sums
df_sales_order = df_sales_order[~df_sales_order["Component"].str.startswith("total")]

# Convert WO_Number to string for merging
df_sales_order["WO_Number"] = df_sales_order["WO_Number"].astype(str).str.strip()

# 🛠 Remove "SO-" prefix from WO_Number in df_sales_order
df_sales_order["WO_Number"] = df_sales_order["WO_Number"].str.replace(r'^SO-', '', regex=True)

# Merge sales orders with API data to add "Picked" column
df_sales_order = df_sales_order.merge(
    word_files_df[['WO_Number', 'status']],  
    on="WO_Number",
    how="left"
)

# Rename status column to "Picked" and replace NaN values with "No"
df_sales_order.rename(columns={"status": "Picked"}, inplace=True)
df_sales_order["Picked"] = df_sales_order["Picked"].fillna("No")

# Filter only picked WOs before grouping
df_sales_order_filtered = df_sales_order[df_sales_order["Picked"] == "Picked"]

# Aggregate Picked Quantities per Component (Only for Picked WOs)
picked_parts = (
    df_sales_order_filtered.groupby("Component")["Required_Qty"]
    .sum()
    .reset_index()
    .rename(columns={"Component": "Part_Number", "Required_Qty": "Picked"})
)


# 🏗 Load Inventory File
try:
    inventory_df = pd.read_csv(warehouse_inventory_file, encoding="ISO-8859-1")
except UnicodeDecodeError:
    inventory_df = pd.read_csv(warehouse_inventory_file, encoding="latin1")

# Standardize part numbers for merging
inventory_df.rename(columns={'Unnamed: 0': 'Part_Number', 'OnHandQty': 'On Hand'}, inplace=True)
inventory_df["Part_Number"] = inventory_df["Part_Number"].astype(str).str.strip().str.lower()

# Merge with Inventory
final_inventory_df = inventory_df.merge(picked_parts, on="Part_Number", how="left")

# Fill missing Picked values with 0
final_inventory_df["Picked"] = final_inventory_df["Picked"].fillna(0)

final_inventory_df['Net Available'] = final_inventory_df["On Hand"] - final_inventory_df["Picked"]


# Read the CSV files with encoding handling
try:
    sales_orders = pd.read_csv(sales_order_file, encoding="ISO-8859-1")
except UnicodeDecodeError:
    sales_orders = pd.read_csv(sales_order_file, encoding="latin1")

try:
    warehouse_inventory = pd.read_csv(warehouse_inventory_file, encoding="ISO-8859-1")
except UnicodeDecodeError:
    warehouse_inventory = pd.read_csv(warehouse_inventory_file, encoding="latin1")

# Rename columns for consistency
sales_orders.rename(columns={sales_orders.columns[0]: 'Component', 'Num': 'WO_Number', 'Qty': 'Required_Qty'}, inplace=True)
warehouse_inventory.rename(columns={'Unnamed: 0': 'Part_Number', 'On Hand': 'Stock_Available'}, inplace=True)

# Fill down the WO_Number to associate components with the correct WO
sales_orders['WO_Number'] = sales_orders['WO_Number'].ffill()

# Assign components based on headers
current_component = None
for idx, row in sales_orders.iterrows():
    if pd.notna(row['Component']) and not row['Component'].startswith("Total "):
        current_component = row['Component']  # Assign new component
    elif pd.notna(row['Component']) and row['Component'].startswith("Total "):
        current_component = None  # Stop forward-filling
    elif current_component:
        sales_orders.at[idx, 'Component'] = current_component  # Assign component name

# Remove "Total" rows as they are no longer needed
sales_orders = sales_orders[~sales_orders['Component'].str.startswith("Total ", na=False)]

# Remove "Forwarding Charge" as it's just shipping cost
sales_orders = sales_orders[~sales_orders['Component'].str.contains("Forwarding Charge", na=False)]

# Standardize names for merging
sales_orders['Component'] = sales_orders['Component'].str.strip().str.lower()
warehouse_inventory['Part_Number'] = warehouse_inventory['Part_Number'].str.strip().str.lower()

# Merge sales orders with inventory (left join to keep all WOs)
structured_df = sales_orders.merge(
    warehouse_inventory, how="left", left_on="Component", right_on="Part_Number"
)

# Remove rows where Required_Qty is NaN
structured_df = structured_df.dropna(subset=['Required_Qty'])


# Determine component status
structured_df['Component_Status'] = structured_df.apply(
    lambda row: "Available" if row['Stock_Available'] >= row['Required_Qty'] else "Shortage",
    axis=1
)

# Identify missing quantities
structured_df['Missing_Qty'] = structured_df.apply(
    lambda row: max(0, row['Required_Qty'] - row['Stock_Available']) if row['Component_Status'] == "Shortage" else 0,
    axis=1
)

# Create an ERP-style hierarchical structure
erp_data = []

for wo_number in structured_df['WO_Number'].unique():
    # Add the main WO entry
    erp_data.append({
        'WO_Number': wo_number,
        'Component': f"Work Order {wo_number}",
        'Required_Qty': '',
        'Stock_Available': '',
        'Component_Status': '',
        'Missing_Qty': ''
    })
    
    # Get the components only for this WO
    wo_components = structured_df[structured_df['WO_Number'] == wo_number]
    
    for _, row in wo_components.iterrows():
        if pd.notna(row['Component']):  # Ensure component is not NaN
            erp_data.append({
                'WO_Number': '',
                'Component': f"  └ {row['Component']}",
                'Required_Qty': row['Required_Qty'],
                'Stock_Available': row['Stock_Available'],
                'Component_Status': row['Component_Status'],
                'Missing_Qty': row['Missing_Qty']
            })

# Convert to DataFrame
erp_display_df = pd.DataFrame(erp_data)

erp_display_df['WO_Number'] = erp_display_df['WO_Number'].ffill()


# Identify rows that are Work Orders (do NOT forward-fill these)
work_order_rows = erp_display_df['Component'].str.contains("Work Order", na=False)

# Forward-fill WO_Number **only for component rows**
erp_display_df['WO_Number'] = erp_display_df['WO_Number'].mask(~work_order_rows).ffill()

# Ensure Component formatting matches for merging
erp_display_df["Component_Cleaned"] = erp_display_df["Component"].str.replace(r'^[^a-zA-Z0-9]+', '', regex=True)
final_inventory_df["Part_Number"] = final_inventory_df["Part_Number"].astype(str).str.strip().str.lower()

# Merge `Picked` and `Net Available` while preserving hierarchical format
merged_df = erp_display_df.merge(
    final_inventory_df[["Part_Number", "Picked", "Net Available"]],
    left_on="Component_Cleaned",
    right_on="Part_Number",
    how="left"
)

# Drop the temporary "Component_Cleaned" and "Part_Number" columns
merged_df.drop(columns=["Component_Cleaned", "Part_Number"], inplace=True)

# Fill missing 'Picked' and 'Net Available' values with 0
merged_df["Picked"] = merged_df["Picked"].fillna(0)
merged_df["Net Available"] = merged_df["Net Available"].fillna(0)

# Ensure Work Order headers do not get Picked or Net Available values
work_order_rows = merged_df['Component'].str.contains("Work Order", case=False, na=False)
merged_df.loc[work_order_rows, ["Picked", "Net Available"]] = ""

# 🖥 Flask Routes
@app.route('/', methods=['GET', 'POST'])
def index():
    search_query = request.form.get('search_query', '').strip()

    # Query Database for PDFs & Word Logs
    pdf_results = PDFFileLog.query.filter(PDFFileLog.file_name.ilike(f"%{search_query}%")).all()
    word_results = WordFileLog.query.filter(WordFileLog.file_name.ilike(f"%{search_query}%")).all()

    # Extract Product Details
    word_data = []
    for word in word_results:
        if word.product_details:
            try:
                product_details_list = json.loads(word.product_details) if isinstance(word.product_details, str) else word.product_details
                for entry in product_details_list:
                    word_data.append({
                        "file_name": word.file_name,
                        "file_path": word.file_path,
                        "product_number": entry.get("product_number", "N/A"),
                        "qty": entry.get("qty", "N/A"),
                        "sn": entry.get("sn", "N/A"),
                        "notes": entry.get("notes", "N/A"),
                    })
            except Exception as e:
                logging.error(f"Error processing product details for {word.file_name}: {str(e)}")

    # **🚀 Ensure filtered_inventory is always initialized**
    if search_query:
        filtered_inventory = merged_df[merged_df["WO_Number"].astype(str).str.contains(search_query, na=False)]
    else:
        filtered_inventory = merged_df  # Show all data if no search query



    return render_template_string("""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Work Order Viewer</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <script src="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/js/all.min.js"></script>
        <style>
            .container { margin-top: 30px; }
            .card-header { font-weight: bold; }
            .file-link { text-decoration: none; font-weight: bold; }
            .table th, .table td { vertical-align: middle; }
            .dataframe { width: 100%; margin-top: 20px; }
            .dataframe thead th { background-color: #343a40; color: white; text-align: right; }
            .dataframe tbody tr:nth-child(odd) { background-color: #f8f9fa; }
            .dataframe tbody tr:nth-child(even) { background-color: #ffffff; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1 class="text-center mb-4">📂 Work Order Viewer</h1>

            <!-- Search Form -->
            <form method="post" class="mb-4">
                <div class="input-group">
                    <input type="text" class="form-control" name="search_query" placeholder="Search Work Order ID" value="{{ search_query }}">
                    <button class="btn btn-primary" type="submit"><i class="fas fa-search"></i> Search</button>
                </div>
            </form>

            <!-- PDF Files -->
            <div class="card mb-4">
                <div class="card-header text-white bg-danger text-start">
                    <i class="fas fa-file-pdf"></i> PDF Files
                </div>
                <div class="card-body">
                    {% if pdf_results %}
                        <ul class="list-group">
                            {% for pdf in pdf_results %}
                                <li class="list-group-item">
                                    <a href="{{ url_for('view_file', file_path=pdf.file_path) }}" target="_blank" class="file-link">
                                        📄 {{ pdf.file_name }}
                                    </a>
                                </li>
                            {% endfor %}
                        </ul>
                    {% else %}
                        <p class="text-muted">No PDF files found.</p>
                    {% endif %}
                </div>
            </div>

                                  
            <!-- Inventory Status Table -->
            {% if not filtered_inventory.empty %}
            <div class="card mb-4">
                <div class="card-header text-white bg-secondary" style="text-align: left;">
                    <i class="fas fa-boxes"></i> Inventory Status
                </div>
                <div class="card-body p-0">
                    <table class="dataframe table table-bordered m-0">
                        <thead>
                            <tr>
                                <th style="text-align: left;">WO_Number</th>
                                <th style="text-align: left;">Component</th>
                                <th style="text-align: left;">Required_Qty</th>
                                <th style="text-align: left;">Stock_Available</th>
                                <th style="text-align: left;">Component_Status</th>
                                <th style="text-align: left;">Missing_Qty</th>
                                <th style="text-align: left;">Picked</th>
                                <th style="text-align: left;">Net Available</th>
                            </tr>
                        </thead>

                        <tbody>
                            {% for idx, entry in filtered_inventory.iterrows() %}
                            <tr>
                                <td>{{ entry.WO_Number }}</td>
                                <td>{{ entry.Component }}</td>
                                <td>{{ entry.Required_Qty }}</td>
                                <td>{{ entry.Stock_Available if entry.Stock_Available is not none else 'NaN' }}</td>
                                <td>{{ entry.Component_Status }}</td>
                                <td>{{ entry.Missing_Qty }}</td>
                                <td>{{ entry.Picked }}</td>
                                <td>{{ entry['Net Available'] }}</td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
            {% else %}
                <p class="text-muted">No inventory data found for this Work Order.</p>
            {% endif %}

            <!-- Word Files Product Details Table -->
            <div class="card mb-4">
                <div class="card-header text-white bg-primary">
                    <i class="fas fa-file-word"></i> Extracted Product Details from Word
                </div>
                <div class="card-body">
                    {% if word_data %}
                        <table class="table table-bordered table-striped">
                            <thead class="table-dark">
                                <tr>
                                    <th>File Name</th>
                                    <th>Product Number</th>
                                    <th>QTY</th>
                                    <th>SN</th>
                                    <th>Notes</th>
                                    <th>File Path</th>
                                </tr>
                            </thead>
                            <tbody>
                                {% for entry in word_data %}
                                <tr>
                                    <td>{{ entry.file_name }}</td>
                                    <td>{{ entry.product_number }}</td>
                                    <td>{{ entry.qty }}</td>
                                    <td>{{ entry.sn }}</td>
                                    <td>{{ entry.notes }}</td>
                                    <td>
                                        <a href="{{ entry.file_path }}" target="_blank" class="file-link">
                                            📂 View File
                                        </a>
                                    </td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    {% else %}
                        <p class="text-muted">No extracted product details available from Word files.</p>
                    {% endif %}
                </div>
            </div>
        </div>
    </body>
    </html>
    """, pdf_results=pdf_results, word_data=word_data, search_query=search_query, filtered_inventory=filtered_inventory)


@app.route('/view_file/')
def view_file():
    file_path = request.args.get('file_path')
    if not file_path or not os.path.exists(file_path):
        abort(404, description="File not found")
    return send_file(file_path, as_attachment=False)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5006)


















