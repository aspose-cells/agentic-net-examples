// Title: Group a date row field by month and year in an Aspose.Cells PivotTable using C#
// AI Prompts: Generate C# code that creates a PivotTable from a date column and groups the row field by months and years with Aspose.Cells. | Show how to use PivotField.GroupBy to apply month and year grouping to a date field in an existing Aspose.Cells PivotTable. | Demonstrate changing the grouping interval to 1 month and 1 year for a date field in a PivotTable with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# group pivot table date field by month and year | example of PivotField.GroupBy months years Aspose.Cells | how to set month/year grouping on date rows in an Aspose.Cells pivot table | C# pivot table date hierarchy grouping Aspose.Cells tutorial
// Tags: Aspose.Cells PivotTable date grouping month year | C# PivotField GroupBy months years | Aspose.Cells create pivot table with date hierarchy | GroupBy method for date fields in Aspose.Cells | PivotTable grouping interval C# Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a workbook, adds date and sales data, builds a PivotTable, then uses PivotField.GroupBy to group the Date row field by months and years (1‑month/1‑year intervals), refreshes the pivot, and saves the file as GroupedByMonthYear.xlsx.
class GroupDateFieldsByMonthYear
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header row
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Sales");

        // Populate sample data spanning several months
        sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 15));
        sheet.Cells["B2"].PutValue(1200);
        sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 10));
        sheet.Cells["B3"].PutValue(1500);
        sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 5));
        sheet.Cells["B4"].PutValue(1800);
        sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 20));
        sheet.Cells["B5"].PutValue(2100);
        sheet.Cells["A6"].PutValue(new DateTime(2023, 5, 25));
        sheet.Cells["B6"].PutValue(2400);

        // Create a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add Date field to rows and Sales field to data area
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Retrieve the date field from the row area
        PivotField dateField = pivot.RowFields[0];

        // Define the grouping range (full year)
        DateTime start = new DateTime(2023, 1, 1);
        DateTime end = new DateTime(2023, 12, 31);

        // Group the date field by Months and Years
        dateField.GroupBy(start, end,
            new PivotGroupByType[] { PivotGroupByType.Months, PivotGroupByType.Years },
            1,          // interval (1 month / 1 year)
            false);     // do not create a new field for the first group type

        // Apply grouping
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook
        workbook.Save("GroupedByMonthYear.xlsx");
    }
}
