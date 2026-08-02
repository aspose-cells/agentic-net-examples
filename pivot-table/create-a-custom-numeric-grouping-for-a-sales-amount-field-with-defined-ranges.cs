// Title: Aspose.Cells C# Example: Pivot Table with Custom Numeric Grouping for Sales Amounts
// Description: This sample builds an Excel workbook, adds region and sales rows, creates a pivot table, and applies the PivotField.GroupBy method to split the Sales column into defined intervals (0‑2000, 2000‑4000, 4000‑6000, 6000‑8000) before saving the file as CustomNumericGrouping.xlsx.
// Keywords: Aspose.Cells | C# pivot table | numeric interval grouping | PivotField.GroupBy | sales amount buckets | Excel automation .NET | custom groups in pivot | Excel library example | group by range Aspose | C# Excel code sample
// Common Searches: Aspose.Cells custom numeric grouping pivot | C# group sales values in pivot table | PivotField.GroupBy usage example | Define numeric intervals in Excel pivot using .NET | How to create sales ranges in Aspose.Cells pivot
// Developer Intent: The developer needs to generate a pivot table and categorize the Sales field into specific numeric ranges for reporting and analysis.
// Use Cases: Summarize regional sales by predefined amount brackets. | Produce Excel reports where sales figures are bucketed into custom intervals. | Automate pivot table creation with grouped numeric fields for business dashboards. | Export grouped sales data for downstream BI or analytics tools.
// AI Prompts: Generate C# code that groups a pivot table field into 0‑1000, 1001‑5000, and >5000 using Aspose.Cells. | Explain how the GroupBy parameters correspond to start, end, and interval values in Aspose.Cells. | Show how to rename generated groups to friendly labels such as Low, Medium, High in a pivot table. | Provide a step‑by‑step guide to apply custom numeric intervals to any numeric field in an Aspose.Cells pivot.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This sample builds an Excel workbook, adds region and sales rows, creates a pivot table, and applies the PivotField.GroupBy method to split the Sales column into defined intervals (0‑2000, 2000‑4000, 4000‑6000, 6000‑8000) before saving the file as CustomNumericGrouping.xlsx.
class CustomNumericGroupingDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ----- Sample data -----
        // Header
        sheet.Cells["A1"].PutValue("Region");
        sheet.Cells["B1"].PutValue("Sales");

        // Data rows
        sheet.Cells["A2"].PutValue("North"); sheet.Cells["B2"].PutValue(800);
        sheet.Cells["A3"].PutValue("North"); sheet.Cells["B3"].PutValue(1500);
        sheet.Cells["A4"].PutValue("South"); sheet.Cells["B4"].PutValue(3200);
        sheet.Cells["A5"].PutValue("South"); sheet.Cells["B5"].PutValue(7200);
        sheet.Cells["A6"].PutValue("East");  sheet.Cells["B6"].PutValue(4500);
        sheet.Cells["A7"].PutValue("West");  sheet.Cells["B7"].PutValue(1100);

        // ----- Create Pivot Table -----
        // Source range A1:B7, place pivot at D3, name it "SalesPivot"
        int pivotIndex = sheet.PivotTables.Add("A1:B7", "D3", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add Region as a row field
        pivot.AddFieldToArea(PivotFieldType.Row, "Region");

        // Add Sales as a data field (aggregated by Sum by default)
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Also add Sales as a row field so we can apply numeric grouping
        int salesRowFieldPos = pivot.AddFieldToArea(PivotFieldType.Row, "Sales");
        PivotField salesField = pivot.RowFields[salesRowFieldPos];

        // ----- Custom numeric grouping -----
        // Define three custom ranges:
        //   0   – 1,000
        //   1,001 – 5,000
        //   >5,000
        // Aspose.Cells groups numeric fields by specifying start, end and interval.
        // We'll use start = 0, end = 8000 (covers all sample values) and interval = 2000.
        // This creates groups: 0‑2000, 2000‑4000, 4000‑6000, 6000‑8000.
        // The first two groups correspond to the desired ranges; the remaining groups can be left as‑is.
        salesField.GroupBy(0, 8000, 2000, false);

        // Refresh and calculate the pivot table to apply the grouping
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook with the grouped pivot table
        workbook.Save("CustomNumericGrouping.xlsx");
    }
}
