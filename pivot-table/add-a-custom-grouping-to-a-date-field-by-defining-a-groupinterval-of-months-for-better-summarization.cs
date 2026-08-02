// Title: Group a Date Field by Months in an Aspose.Cells PivotTable (C#)
// Description: Creates a workbook, adds sample dates and sales, builds a pivot table, then groups the Date row field into 1‑month intervals using PivotField.GroupBy without creating a new field. The pivot is refreshed, calculated, and saved as GroupedByMonthsPivot.xlsx.
// Keywords: Aspose.Cells | C# PivotTable | GroupBy months | date field grouping | custom group interval | pivot table summarization | Aspose.Cells GroupBy overload | monthly sales pivot
// Common Searches: Aspose.Cells group date field by month | C# PivotTable GroupBy example | how to group dates in Aspose.Cells pivot | monthly grouping without new field Aspose.Cells | custom group interval PivotTable C#
// Developer Intent: Apply month‑level grouping to a Date field in an Aspose.Cells pivot table.
// Use Cases: Convert daily transaction data into monthly totals for reporting. | Create a compact sales dashboard that aggregates dates by month. | Generate periodic financial summaries without adding extra worksheet columns.
// AI Prompts: Show how to change the grouping interval to quarters using Aspose.Cells GroupBy. | Provide code to apply month grouping to multiple date fields in the same pivot table. | Explain how to extract the displayed month labels after grouping with GroupBy.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook, adds sample dates and sales, builds a pivot table, then groups the Date row field into 1‑month intervals using PivotField.GroupBy without creating a new field. The pivot is refreshed, calculated, and saved as GroupedByMonthsPivot.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ---------- Populate sample data ----------
        // Header row
        sheet.Cells["A1"].Value = "Date";
        sheet.Cells["B1"].Value = "Sales";

        // Sample dates and corresponding sales values
        DateTime[] dates = {
            new DateTime(2023, 1, 5),
            new DateTime(2023, 1, 15),
            new DateTime(2023, 2, 10),
            new DateTime(2023, 2, 20),
            new DateTime(2023, 3, 5),
            new DateTime(2023, 3, 25)
        };
        int[] sales = { 1000, 1500, 2000, 2500, 3000, 3500 };

        // Write data to worksheet
        for (int i = 0; i < dates.Length; i++)
        {
            sheet.Cells[i + 2, 0].Value = dates[i];   // Column A (Date)
            sheet.Cells[i + 2, 1].Value = sales[i];   // Column B (Sales)
        }

        // ---------- Create Pivot Table ----------
        // Data range A1:B7, place pivot at E3, name it "SalesPivot"
        int pivotIndex = sheet.PivotTables.Add("A1:B7", "E3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add Date field to Row area and Sales field to Data area
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Get the PivotField representing the Date column
        PivotField dateField = pivotTable.RowFields[0];

        // ---------- Group Date Field by Months ----------
        // Define grouping interval: 1 month, no new field is created
        DateTime startDate = new DateTime(2023, 1, 1);
        DateTime endDate   = new DateTime(2023, 12, 31);
        PivotGroupByType[] groupTypes = new PivotGroupByType[] { PivotGroupByType.Months };
        double interval = 1;          // 1 month
        bool firstAsNewField = false; // keep grouping in the same field

        // Apply grouping using the appropriate GroupBy overload
        dateField.GroupBy(startDate, endDate, groupTypes, interval, firstAsNewField);

        // Refresh and calculate the pivot table to reflect grouping
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // ---------- Save the workbook ----------
        workbook.Save("GroupedByMonthsPivot.xlsx");
    }
}
