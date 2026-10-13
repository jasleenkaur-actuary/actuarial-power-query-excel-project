# Actuarial Data Preparation using Power Query & Excel

Two practice projects that turn raw insurance data into clean, model-ready tables using Power Query (M language) in Excel.

**Note:** These are practice projects based on Power Query exercises, using sample data.

## Tools Used
- Microsoft Excel
- Power Query (M language)

## Projects

| # | Project | Data | Folder |
|---|---|---|---|
| 1 | Annuity Data Preparation | 50,000 annuity policies | [01_Annuity_Data_Prep](01_Annuity_Data_Prep) |
| 2 | GI Data Transformation | General insurance loss data | [02_GI_Data_Transformation](02_GI_Data_Transformation) |

---

## Case Study 1: Annuity Data Preparation

**Problem:** Raw annuity policy data had duplicate Policy IDs, text dates and inconsistent column names, so it could not be used in a financial model.

**What I did:**
- Converted date columns to Date type
- Removed duplicate Policy IDs, keeping the latest Update Date
- Renamed and reordered columns to the model specification
- Calculated age at the 31 March 2026 valuation date
- Calculated payment per period from the payment frequency
- Assigned a risk band by age (50-60 Low, 61-75 Medium, 76+ High)
- Sorted the final table by PolicyID

**Result:** 50,000 clean, model-ready policies.

**Example: risk band logic**
```
AddedRiskBand = Table.AddColumn(AddedPayPerPeriod, "RiskBand", each
    if [AgeAtValuationDate] >= 76 then "High"
    else if [AgeAtValuationDate] >= 61 then "Medium"
    else if [AgeAtValuationDate] >= 50 then "Low"
    else null),
```

**Code:** [annuity_formatted_data.pq](01_Annuity_Data_Prep/annuity_formatted_data.pq)

---

## Case Study 2: General Insurance (GI) Data Transformation

**Problem:** Raw GI loss data was stored in a wide layout (Dev_12, Dev_24, Dev_36 as separate columns), which is hard to analyse and filter.

**What I did:**
- Set the correct data types
- Renamed columns to readable names
- Unpivoted the development period columns into rows
- Converted "Dev_12" into the number 12
- Built a second query that filters the data by a Line of Business and Region selected in the sheet

**Result:** A clean long-format table (Line of Business, Accident Year, Region, Development Period, Incurred Loss Amount) plus a filtered view that updates when the selection changes.

**Example: unpivot step**
```
Unpivoted = Table.UnpivotOtherColumns(
    RenamedCols,
    {"Line of Business", "Accident Year", "Region"},
    "Development Period",
    "Incurred Loss Amount"
),
```

**Code:**
- [gi_formatted_data.pq](02_GI_Data_Transformation/gi_formatted_data.pq)
- [gi_filtered_view.pq](02_GI_Data_Transformation/gi_filtered_view.pq)

---

## Skills Shown
- Data cleaning and standardisation in Power Query
- Removing duplicates while keeping the latest record
- Unpivoting data for actuarial analysis
- Writing conditional logic and calculated columns in M
- Building a filter driven by cells in the sheet

## Folder Structure
```
actuarial-power-query-excel-project/
├── README.md
├── 01_Annuity_Data_Prep/
│   ├── Annuity_Financial_Model_Data_Preparation.xlsx
│   └── annuity_formatted_data.pq
└── 02_GI_Data_Transformation/
    ├── Simple_GI_Actuarial_Data.xlsx
    ├── gi_formatted_data.pq
    └── gi_filtered_view.pq
```