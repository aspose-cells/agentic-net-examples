// Title: Group a Pivot Table Date Field by Month and Year using Aspose.Cells for .NET (C#)
// Description: Creates a workbook with sample dates and values, builds a pivot table, places the Date column in rows and the Value column in data, then groups the Date field into month‑ and year‑level hierarchies via PivotField.GroupBy, refreshes the pivot, and saves the result as PivotGroupedByMonthsYears.xlsx.
// Keywords: Aspose.Cells pivot table month grouping | C# group pivot date by year | Aspose.Cells hierarchical date grouping | PivotField.GroupBy months years .NET | Excel pivot date hierarchy Aspose | US developers Aspose.Cells pivot
// Common Searches: Aspose.Cells group pivot date by month and year C# | How to create month‑year hierarchy in Aspose pivot table | PivotField.GroupBy example Aspose.Cells .NET | Date grouping in Aspose.Cells pivot tables
// Developer Intent: Apply month and year grouping to a pivot table's Date field for hierarchical reporting.
// Use Cases: Monthly and yearly sales aggregation in a financial workbook. | Generating a timeline view of transactions with month‑year drill‑down. | Building a dashboard that summarizes KPI data by month and then by year.
// AI Prompts: Write C# code with Aspose.Cells that groups a pivot table Date field into months and years and saves the workbook. | Explain each parameter of PivotField.GroupBy when grouping dates in Aspose.Cells. | Show how to extend the grouping to include quarters or days alongside months and years in an Aspose.Cells pivot table.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook with sample dates and values, builds a pivot table, places the Date column in rows and the Value column in data, then groups the Date field into month‑ and year‑level hierarchies via PivotField.GroupBy, refreshes the pivot, and saves the result as PivotGroupedByMonthsYears.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Prepare sample data: a Date column and a Value column
        // -------------------------------------------------
        worksheet.Cells["A1"].PutValue("Date");
        worksheet.Cells["B1"].PutValue("Value");

        // Add several dates spanning multiple months
        worksheet.Cells["A2"].PutValue(new DateTime(2023, 1, 5));
        worksheet.Cells["A3"].PutValue(new DateTime(2023, 2, 12));
        worksheet.Cells["A4"].PutValue(new DateTime(2023, 3, 20));
        worksheet.Cells["A5"].PutValue(new DateTime(2023, 4, 8));
        worksheet.Cells["A6"].PutValue(new DateTime(2023, 5, 15));

        // Corresponding numeric values
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["B4"].PutValue(200);
        worksheet.Cells["B5"].PutValue(250);
        worksheet.Cells["B6"].PutValue(300);

        // -------------------------------------------------
        // Create a pivot table based on the data range
        // -------------------------------------------------
        int pivotIndex = worksheet.PivotTables.Add("A1:B6", "E3", "SalesPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add the Date field to the row area and the Value field to the data area
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // column index 0 -> Date
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // column index 1 -> Value

        // -------------------------------------------------
        // Group the Date field by Months and Years
        // -------------------------------------------------
        PivotField dateField = pivotTable.RowFields[0];

        // Define the grouping range (full year) and the desired group types
        DateTime startDate = new DateTime(2023, 1, 1);
        DateTime endDate   = new DateTime(2023, 12, 31);
        PivotGroupByType[] groupTypes = new PivotGroupByType[]
        {
            PivotGroupByType.Months,
            PivotGroupByType.Years
        };

        // Apply grouping; interval is ignored when specific group types are supplied
        dateField.GroupBy(startDate, endDate, groupTypes, 1, false);

        // Refresh and calculate the pivot table to reflect the new grouping
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("PivotGroupedByMonthsYears.xlsx");
    }
}
