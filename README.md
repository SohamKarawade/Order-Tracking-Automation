# 📦 Order Lifecycle Automation

A Python-based automation system that monitors Microsoft Outlook email conversations to track e-commerce orders across the E-commerce, CRM, and Inventory teams. The application automatically extracts order information, tracks the complete order lifecycle, generates Excel reports with dashboards, and enables one-click report sharing through a user-friendly frontend.

---

# 🚀 Overview

Managing customer orders manually through email conversations is time-consuming and prone to errors. This project automates the entire process by continuously monitoring Outlook emails, extracting order information, identifying order progress, generating reports, and providing real-time summaries for management.

The automation runs every **5 minutes**, ensuring reports are always up to date.

---

# 🎯 Business Problem

The operations team manually tracked every order by checking Outlook email conversations.

For each order, they had to:

* Identify new order emails from the E-commerce team.
* Verify whether the CRM team had confirmed the order.
* Check whether the Inventory team had completed quality checks and dispatch.
* Track replacement or cancellation requests.
* Prepare daily summary reports for management.

This process required manually reviewing hundreds of email conversations every day.

---

# ✅ Solution

This project automates the complete workflow.

The system:

* Connects directly to Microsoft Outlook.
* Identifies valid order conversations.
* Extracts customer and order details.
* Tracks CRM and Inventory responses.
* Detects replacement and cancellation requests.
* Automatically calculates order status.
* Generates formatted Excel reports.
* Creates management dashboards.
* Shares reports through email with a single click.

---

# 🔄 Workflow

```text
Customer Order
      │
      ▼
E-commerce Team sends Order Mail
      │
      ▼
Automation detects valid order mail
      │
      ▼
Extracts Order Details
      │
      ▼
Checks CRM Confirmation
      │
      ▼
Checks Inventory Dispatch
      │
      ▼
Checks Replacement / Cancellation
      │
      ▼
Generates Excel Report
      │
      ▼
Creates Summary Dashboard
      │
      ▼
Shares Report via Email
```

---

# ✨ Features

## Outlook Automation

* Reads Outlook emails using Python
* Processes only valid order conversations
* Filters emails using Conversation ID, sender, and subject

## Order Data Extraction

Automatically extracts:

* Conversation ID
* Order ID
* Shopify ID
* Customer Name
* Product Name
* SKU Code
* Quantity
* Customer Contact Number
* Order Date
* E-commerce Mail Date
* Email Subject

## CRM Tracking

Tracks:

* CRM mail received (Yes/No)
* CRM response date
* Replacement requests
* Cancellation requests

## Inventory Tracking

Tracks:

* Inventory mail received (Yes/No)
* Inventory mail date

## Intelligent Status Engine

Automatically determines order status:

| Condition                           | Status                  |
| ----------------------------------- | ----------------------- |
| CRM not received                    | Pending                 |
| CRM received, Inventory pending     | Confirmed               |
| Inventory completed                 | Delivered               |
| Replacement / Cancellation detected | Replacement / Cancelled |

Conditional formatting is automatically applied to the Excel report.

---

# 📊 Excel Reports

The automation generates a workbook containing:

### Order Data

Complete tracking information for every order.

### Summary Dashboard

Automatically compares:

* Today vs Yesterday
* This Week vs Last Week
* This Month vs Last Month

Metrics include:

* Orders Received
* Pending Orders
* Confirmed Orders
* Delivered Orders

### Additional Sheets

Separate worksheets are generated for:

* Today's Orders
* Yesterday's Orders
* This Week's Orders
* Last Week's Orders
* This Month's Orders
* Last Month's Orders

---

# 🖥 Frontend Features

A simple frontend allows users to perform common tasks without running Python scripts manually.

### Generate Report

* Generates a fresh report using the latest Outlook data.

### Share Report

* Generates the latest report.
* Attaches the Excel file.
* Creates a preformatted email.
* Shares the report with management.

---

# ⚙️ Technologies Used

* Python
* Microsoft Outlook COM Automation (win32com)
* FastAPI
* HTML
* CSS
* JavaScript
* Pandas
* OpenPyXL
* Regular Expressions (Regex)

---

# 📁 Project Structure

```
Order-Lifecycle-Automation/
│
├── backend/
├── frontend/
├── scheduler/
├── reports/
├── screenshots/
├── sample_output/
├── README.md
├── requirements.txt
└── LICENSE
```

---

# ▶️ Getting Started

## Clone the repository

```bash
git clone https://github.com/yourusername/Order-Lifecycle-Automation.git
```

## Install dependencies

```bash
pip install -r requirements.txt
```

## Configure

Update the following before running:

* Outlook email account
* Sender email addresses
* Output folder
* Report recipients

---

# 📈 Business Impact

* Eliminated manual tracking of order emails.
* Reduced report preparation time through automation.
* Improved reporting accuracy.
* Enabled near real-time order visibility.
* Standardized operational reporting.
* Simplified report generation and sharing.

---

# 🔮 Future Enhancements

* Shopify API integration
* Database support (PostgreSQL / SQL Server)
* Power BI dashboard integration
* Docker deployment
* Cloud hosting
* Multi-user authentication
* Real-time notifications
* Web-based dashboard

---

# 📸 Screenshots

Add screenshots here:

* Frontend Dashboard
* Excel Order Data
* Summary Dashboard
* Generated Email
* Project Workflow

---

# 📄 License

This project is intended for educational and portfolio purposes. Any confidential company information, customer data, email addresses, and business-specific details have been removed or anonymized.

---

## 👨‍💻 Author

**Soham Karawade**

If you found this project useful, feel free to ⭐ the repository.
