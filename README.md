# 📊 Customer Support Ticket Analysis | Microsoft Excel

## 📌 Project Overview

The **Customer Support Ticket Analysis** project uses Microsoft Excel to analyze customer support tickets, evaluate support team performance, and monitor key customer service metrics through an interactive-style dashboard.

The project focuses on understanding ticket distribution, response times, resolution times, ticket priorities, customer satisfaction, and ticket status to help identify opportunities for improving customer support operations.

## Objectives

- Analyze the total volume of customer support tickets.
- Identify the most common ticket categories and customer support channels.
- Evaluate ticket distribution by priority and status.
- Measure average first response time and resolution time.
- Track customer satisfaction using Customer Satisfaction Score (CSAT).
- Monitor resolution rates and high-priority tickets.
- Present key performance indicators (KPIs) in an Excel dashboard.

## Tools & Technologies

- **Microsoft Excel**
- Pivot Tables and Pivot Analysis
- Excel Formulas and Functions
- KPI Reporting and Dashboard Design
- Data Analysis and Visualization

## Project Structure

| Worksheet | Description |
|---|---|
| `Ticket Data` | Contains customer support ticket records, including ticket ID, opened date, category, priority, channel, assigned agent, status, response time, resolution time, and CSAT score. |
| `Pivot Analysis` | Summarizes ticket distribution across categories, priorities, channels, and support agents. |
| `Dashboard` | Displays key performance indicators to help monitor customer support performance. |

## Key Performance Indicators (KPIs)

The dashboard includes the following metrics:

- **Total Tickets:** Total number of recorded support tickets.
- **Resolved Tickets:** Number of tickets marked as resolved.
- **Average CSAT:** Average customer satisfaction score.
- **Average Response Time:** Average time taken to provide the first response, measured in minutes.
- **Average Resolution Time:** Average time taken to resolve tickets, measured in hours.
- **Resolution Rate:** Percentage of tickets resolved.
- **High-Priority Tickets:** Number of tickets classified as high priority.
- **Critical-Priority Tickets:** Number of tickets classified as critical priority.
- **High/Critical Resolved Tickets:** Number of resolved tickets belonging to the high and critical priority groups.
- **Average High-Priority Response Time:** Average first response time for high-priority tickets.
- **Average Critical-Priority Resolution Time:** Average resolution time for critical-priority tickets.

## Excel Formulas Used

The project uses Excel functions to calculate and monitor customer support metrics.

| Function | Purpose |
|---|---|
| `COUNTA()` | Counts recorded ticket IDs. |
| `COUNTIF()` | Counts tickets matching a specific condition, such as resolved status or ticket priority. |
| `COUNTIFS()` | Counts tickets that meet multiple conditions. |
| `AVERAGE()` | Calculates average customer satisfaction and response or resolution times. |
| `AVERAGEIF()` | Calculates average metrics for a particular ticket priority. |

### Example Formulas

**Total Tickets**
```excel
=COUNTA(Tickets[TicketID])
```

**Resolved Tickets**
```excel
=COUNTIF(Tickets[Status],"Resolved")
```

**Average Customer Satisfaction**
```excel
=AVERAGE(Tickets[CSAT])
```

**Resolution Rate**
```excel
=COUNTIF(Tickets[Status],"Resolved")/COUNTA(Tickets[TicketID])
```

**Average Response Time for High-Priority Tickets**
```excel
=AVERAGEIF(Tickets[Priority],"High",Tickets[FirstResponse_Min])
```

## Analysis Performed

The project examines customer support performance across several dimensions:

1. **Ticket Category Analysis:** Compares ticket volumes across Account, Delivery, Payment, Refund, and Technical categories.
2. **Priority Analysis:** Examines ticket distribution across Low, Medium, High, and Critical priorities.
3. **Channel Analysis:** Compares tickets received through Chat, Email, and Phone.
4. **Agent Analysis:** Organizes ticket data by assigned support agent.
5. **Resolution Analysis:** Tracks resolved, pending, and escalated tickets.
6. **Response and Resolution Analysis:** Evaluates average first response and resolution times.
7. **Customer Satisfaction Analysis:** Uses CSAT scores to assess customer experience.

## Business Value

This project demonstrates how Excel can support data-driven customer service decisions by helping teams:

- Identify categories generating higher ticket volumes.
- Understand workload distribution across support channels and agents.
- Monitor response and resolution efficiency.
- Identify high-priority and critical tickets that may require attention.
- Evaluate customer satisfaction.
- Track support performance through a centralized dashboard.

## How to Use

1. Download the Excel workbook from this repository.
2. Open the file using Microsoft Excel.
3. Navigate to the **Ticket Data** sheet to explore the underlying records.
4. Review the **Pivot Analysis** sheet to examine ticket distribution.
5. Open the **Dashboard** sheet to view the key performance indicators.
6. Modify the data to explore how the analysis changes. If you add records, ensure the Excel table and any Pivot Tables are updated or refreshed as needed.

## Skills Demonstrated

- Data cleaning and organization
- Exploratory data analysis
- Pivot Table analysis
- Excel formulas and conditional calculations
- KPI development
- Dashboard reporting
- Business problem-solving
- Customer support performance analysis

## Dataset Information

The workbook contains **60 sample customer support ticket records** across multiple categories, priorities, channels, agents, and ticket statuses.

The dataset is intended for learning, demonstration, and portfolio purposes and should not be interpreted as actual business performance data.

## Author

**Vaibhavi Suvarna**


---

⭐ If you find this project useful, consider giving the repository a star!
