// Title: Group Numeric Pivot Field Values into Custom Ranges with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a PivotTable, place a numeric "Amount" field in the row area, and use PivotField.GroupBy to segment the values into custom intervals (0‑10, 10‑20, 20‑30, 30‑40, 40‑60). The example refreshes the pivot cache, recalculates the table, and saves the result as an Excel file.
// Keywords: Aspose.Cells | C# PivotTable | numeric grouping | custom range grouping | PivotField.GroupBy | .NET Excel API | group rows by interval | Excel pivot custom buckets
// Common Searches: Aspose.Cells group numeric pivot values | PivotTable custom range grouping .NET | How to use PivotField.GroupBy in C# | Create numeric buckets in Aspose.Cells pivot | Aspose.Cells interval grouping example
// Developer Intent: Create a PivotTable and group a numeric row field into predefined ranges using Aspose.Cells for .NET.
// Use Cases: Break down sales amounts into revenue brackets for management reports. | Classify ages into demographic groups within a pivot for market analysis. | Bucket financial figures (e.g., profit ranges) for dashboard visualizations.
// AI Prompts: Show C# code that uses Aspose.Cells PivotField.GroupBy to group numeric values into custom intervals and refreshes the pivot. | Explain each parameter of the PivotField.GroupBy method and how to apply it without creating a new field. | Generate an Aspose.Cells example that groups a numeric column into ranges 0‑5, 5‑15, 15‑25 in a PivotTable.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook, add a PivotTable, place a numeric "Amount" field in the row area, and use PivotField.GroupBy to segment the values into custom intervals (0‑10, 10‑20, 20‑30, 30‑40, 40‑60). The example refreshes the pivot cache, recalculates the table, and saves the result as an Excel file.
class PivotNumericCustomRangeGrouping
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: a category column and a numeric amount column
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["B1"].Value = "Amount";

        sheet.Cells["A2"].Value = "A";
        sheet.Cells["B2"].Value = 5;
        sheet.Cells["A3"].Value = "B";
        sheet.Cells["B3"].Value = 12;
        sheet.Cells["A4"].Value = "C";
        sheet.Cells["B4"].Value = 27;
        sheet.Cells["A5"].Value = "D";
        sheet.Cells["B5"].Value = 33;
        sheet.Cells["A6"].Value = "E";
        sheet.Cells["B6"].Value = 48;
        sheet.Cells["A7"].Value = "F";
        sheet.Cells["B7"].Value = 55;

        // Create a pivot table based on the data range A1:B7, place it at D3
        int pivotIdx = sheet.PivotTables.Add("A1:B7", "D3", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIdx];

        // Add the numeric field "Amount" to the row area (this will be grouped)
        pivot.AddFieldToArea(PivotFieldType.Row, "Amount");

        // Add the "Category" field to the data area (summarize by count)
        pivot.AddFieldToArea(PivotFieldType.Data, "Category");

        // Access the row field that contains numeric values
        PivotField amountField = pivot.RowFields[0];

        // Group numeric values into custom ranges: 0‑10, 10‑20, 20‑30, 30‑40, 40‑60
        // Parameters: start = 0, end = 60, interval = 10, newField = false (group in place)
        amountField.GroupBy(0, 60, 10, false);

        // Refresh the pivot cache and calculate the pivot table to apply the grouping
        pivot.RefreshData();          // Updated API usage
        pivot.CalculateData();

        // Save the workbook with the grouped pivot table
        workbook.Save("PivotNumericCustomRangeGrouping.xlsx");
    }
}
