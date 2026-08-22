Overview
Developed an end-to-end file share audit analytics solution using VBA, Excel, Power BI, and Windows Security Event Logs (Event ID 5145) to automate the identification, filtering, analysis, and reporting of public network share access across an enterprise healthcare environment. The solution processes raw Windows Security Audit logs, extracts key access metadata, removes non-relevant activity, identifies user access patterns, and generates automated dashboards for compliance, operational monitoring, and investigative reporting.

Key Features
Automated Event Parsing
• Parsed raw Windows Security Audit Event ID 5145 logs.
• Extracted:
  ◦ Security ID
  ◦ User Name
  ◦ Account Domain
  ◦ Share Name
  ◦ Share Path
  ◦ Share Path Name
  ◦ Relative Target Name
  ◦ Source Address
• Data scrubbing:
  • Standardized usernames through:
    ◦ trimming
    ◦ lowercase conversion
    ◦ character cleanup
    • Removed invalid and incomplete records.
    • Normalized share and folder names.
Access Classification
Implemented automated filtering rules to identify:
✅ Public folder activity
✅ Folder access events
✅ Non-E3 user activity
Excluded:
❌ E3 licensed users
❌ File accesses
❌ Blank paths
❌ Non-public share activity
❌ Corrupt audit logs

Reporting & Analytics
Excel Reporting
Generated:
• Detailed audit datasets
• Summarized folder access reports
• Automated PivotTables
Power BI Dashboard
Developed an interactive analytics dashboard featuring:
• Total Access Events
• Unique Users
• Unique Folders Accessed
• Unique Devices
• Top Users by Access Count
• Top Folders by Access Count
• Access Activity Trends
• Dynamic Folder Filtering

Technologies Used
• Microsoft Excel
• VBA (Visual Basic for Applications)
• Power BI
• Power Query
• Windows Security Event Logs
• PivotTables
• DAX
• Data Modeling
• ETL (Extract, Transform, Load)

Impact
• Automated analysis of datasets exceeding 150,000 security audit events.
• Reduced manual audit effort through automated parsing and classification.
• Produced compliance-ready reporting for file share access reviews.
• Enabled trend analysis and user-behavior monitoring through Power BI visualizations
