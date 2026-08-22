File Share Audit Analytics

Overview

Developed an end-to-end **file share audit analytics solution** using **VBA, Excel, Power BI, and Windows Security Event Logs (Event ID 5145)** to automate the identification, filtering, analysis, and reporting of public network share access across an enterprise healthcare environment.

The solution processes raw Windows Security Audit logs, extracts key access metadata, removes non-relevant activity, identifies user access patterns, and generates automated dashboards for **compliance, operational monitoring, and investigative reporting**.

---

Key Features

🔍 Automated Event Parsing

* Parsed raw **Windows Security Audit Event ID 5145** logs.
* Extracted key access metadata, including:

  * Security ID
  * User Name
  * Account Domain
  * Share Name
  * Share Path
  * Share Path Name
  * Relative Target Name
  * Source Address

🧼 Data Scrubbing & Normalization

* Standardized usernames through:

  * Trimming whitespace
  * Lowercase conversion
  * Character cleanup
* Removed invalid and incomplete records.
* Normalized share and folder names.

🏷️ Access Classification

Implemented automated filtering rules to identify:

* ✅ Public folder activity
* ✅ Folder access events
* ✅ Non-E3 user activity

Excluded:

* ❌ E3 licensed users
* ❌ File access events
* ❌ Blank paths
* ❌ Non-public share activity
* ❌ Corrupt or incomplete audit logs

---

📊 Reporting & Analytics

📋 Excel Reporting

Generated automated Excel reports containing:

* Detailed audit datasets
* Summarized folder access reports
* Automated PivotTables
* Filtered and cleaned audit data

🖥️  Power BI Dashboard

Developed an interactive analytics dashboard featuring:

* **Total Access Events**
* **Unique Users**
* **Unique Folders Accessed**
* **Unique Devices**
* **Top Users by Access Count**
* **Top Folders by Access Count**
* **Access Activity Trends**
* **Dynamic Folder Filtering**

---

## Technologies Used

| Technology                      | Purpose                                  |
| ------------------------------- | ---------------------------------------- |
| **Microsoft Excel**             | Data analysis and reporting              |
| **VBA**                         | Automation and event processing          |
| **Power Query**                 | Data transformation and ETL              |
| **Power BI**                    | Interactive dashboards and visualization |
| **DAX**                         | Calculations and analytical measures     |
| **Windows Security Event Logs** | Source audit data                        |
| **PivotTables**                 | Data summarization                       |
| **Data Modeling**               | Structuring analytical datasets          |
| **ETL**                         | Extract, Transform, Load workflow        |

---

Data Processing Workflow

```text
Windows Security Event Logs
            │
            ▼
      Raw Event Data
            │
            ▼
      Event ID 5145 Parsing
            │
            ▼
    Data Cleaning & Normalization
            │
            ▼
      Access Classification
            │
            ▼
      Filtering & Validation
            │
            ▼
      Excel / Power Query
            │
            ▼
       Data Modeling
            │
            ▼
Interactive Analytics Monthly Dashboard with Power BI

```

Impact

* Automated analysis of **150,000+ Windows security audit events**.
* Reduced manual audit effort through automated parsing, filtering, and classification.
* Produced **compliance-ready reporting** for file share access reviews.
* Enabled trend analysis and user-behavior monitoring through Power BI visualizations.
* Created a repeatable workflow for transforming raw security logs into actionable analytics.

---

Skills Demonstrated

* **Data Analysis**
* **ETL & Data Transformation**
* **VBA Automation**
* **Power Query**
* **Power BI**
* **DAX**
* **Data Modeling**
* **Windows Security Event Analysis**
* **Security Audit Log Processing**
* **Dashboard Development**
* **Automation & Reporting**
