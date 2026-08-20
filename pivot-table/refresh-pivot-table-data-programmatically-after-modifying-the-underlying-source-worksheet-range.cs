// Title: Refresh a Pivot Table After Changing Source Data with Aspose.Cells for .NET
// Description: This example creates a workbook, builds a pivot table from a range, updates the source cells, then uses Worksheet.RefreshPivotTables() and PivotTable.CalculateData() to synchronize the pivot report before saving the file.
// Keywords: Aspose.Cells refresh pivot table | Worksheet.RefreshPivotTables | PivotTable.CalculateData | update pivot cache .NET | programmatic pivot refresh | C# Aspose.Cells pivot example
// Common Searches: how to refresh pivot tables in Aspose.Cells | Aspose.Cells refresh pivot after editing cells | C# update pivot cache programmatically | Worksheet.RefreshPivotTables usage | recalculate pivot table Aspose.Cells .NET
// Developer Intent: Synchronize a pivot table with modified source data using Aspose.Cells APIs.
// Use Cases: Automated financial reports where source amounts change and the pivot must show the latest totals. | Dynamic dashboards that add or reclassify rows at runtime; call RefreshPivotTables to keep all pivots consistent. | Batch processing jobs that generate workbooks, modify data, and need an up‑to‑date pivot before distribution.
// AI Prompts: Show C# code that updates source cells and refreshes all pivot tables in a workbook using Aspose.Cells. | Explain the steps to recalculate a single pivot table after changing its source range with Aspose.Cells for .NET. | Compare Worksheet.RefreshPivotTables() and PivotTable.CalculateData() and when each should be used.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This example creates a workbook, builds a pivot table from a range, updates the source cells, then uses Worksheet.RefreshPivotTables() and PivotTable.CalculateData() to synchronize the pivot report before saving the file.
class RefreshPivotDemo
{
    static void Main()
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
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet ws = workbook.Worksheets[0];

        // Populate source data for the pivot table
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["B1"].PutValue("Amount");
        ws.Cells["A2"].PutValue("Food");
        ws.Cells["B2"].PutValue(100);
        ws.Cells["A3"].PutValue("Drink");
        ws.Cells["B3"].PutValue(150);
        ws.Cells["A4"].PutValue("Food");
        ws.Cells["B4"].PutValue(200);

        // Add a pivot table based on the source range
        int ptIndex = ws.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = ws.PivotTables[ptIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Initial calculation to generate the pivot report
        pivotTable.CalculateData();

        // Modify the underlying source data
        ws.Cells["B2"].PutValue(120);   // Change Food amount
        ws.Cells["B3"].PutValue(180);   // Change Drink amount
        ws.Cells["A4"].PutValue("Drink"); // Change category to Drink
        ws.Cells["B4"].PutValue(220);

        // Refresh all pivot tables in the worksheet to reflect the changes
        ws.RefreshPivotTables();

        // Recalculate after refresh (optional, ensures the pivot cache is applied)
        pivotTable.CalculateData();

        // Save the updated workbook
        workbook.Save("RefreshedPivot.xlsx");
        Console.WriteLine("Workbook saved as RefreshedPivot.xlsx");
    }
}
