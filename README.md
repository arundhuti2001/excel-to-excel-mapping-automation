# excel-to-excel-mapping-automation

# AI-Powered Excel Data Automation

## Overview

This project automates the extraction, mapping, validation, and consolidation of data from multiple Excel workbooks into a centralized reporting template. It demonstrates intelligent data processing workflows commonly used in financial reporting, business operations, and enterprise automation.

The solution automatically identifies matching records, creates new entries when required, and updates target datasets while preserving existing workbook structure and formatting.

## Features

- Automated Excel-to-Excel data integration
- Multi-source data consolidation
- Dynamic date matching and validation
- Automatic row creation for new records
- Intelligent data mapping across worksheets
- Preservation of existing template structure
- Automated monthly reporting workflow
- Scalable and reusable automation framework

## Technologies Used

- Python
- OpenPyXL
- Datetime
- Pathlib

## Workflow

1. Load reporting template workbook.
2. Read multiple source Excel files.
3. Parse and validate dates.
4. Match records dynamically.
5. Create missing records automatically.
6. Populate target columns with mapped data.
7. Preserve historical data and formatting.
8. Generate final consolidated report.

## Automation Logic

### Data Mapping
- Source data is extracted from multiple Excel files.
- Records are matched using date-based logic.
- Missing records are automatically inserted.

### Data Population
- Columns B:M populated from source workbook.
- Column N populated from benchmark dataset.
- Column O populated from adjusted index dataset.
- Column P populated from sector index dataset.

### Validation
- Supports multiple date formats.
- Handles Excel serial dates.
- Skips invalid or missing records safely.

## Business Applications

- Financial Reporting Automation
- Index Data Management
- Portfolio Reporting
- Data Consolidation
- ETL Automation
- Business Process Automation
- Operations Reporting

## Benefits

- Reduces manual effort
- Improves data accuracy
- Eliminates repetitive tasks
- Speeds up monthly reporting cycles
- Supports enterprise-scale reporting workflows

## Installation

```bash
pip install openpyxl
```

## Run

```bash
python main.py
```

## Output

The automation generates a consolidated Excel report with mapped and updated data while maintaining the original template structure.
