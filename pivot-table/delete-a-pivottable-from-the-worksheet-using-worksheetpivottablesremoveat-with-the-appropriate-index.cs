// Title: Remove a PivotTable by Index in Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds three pivot tables to the first worksheet, deletes the second table using sheet.PivotTables.RemoveAt(1), and saves the result as DeletedPivotTable.xlsx.
// Keywords: Aspose.Cells | .NET | C# | PivotTable RemoveAt | delete pivot table | remove pivot table by index | Aspose.Cells PivotTables.RemoveAt | programmatic pivot table removal | C# workbook pivot table | Aspose.Cells worksheet pivot tables
// Common Searches: Aspose.Cells remove pivot table by index | C# PivotTables.RemoveAt example | delete specific pivot table Aspose.Cells | how to delete a pivot table in .NET | remove second pivot table Aspose.Cells
// Developer Intent: The developer needs to programmatically delete a particular pivot table from a worksheet using its zero‑based index with Aspose.Cells for .NET.
// Use Cases: Eliminate unwanted pivot tables after generating multiple reports on the same sheet. | Reduce workbook size by removing temporary pivot tables before saving. | Implement user‑driven or rule‑based cleanup of pivot tables in automated reporting pipelines.
// AI Prompts: List all pivot tables in a worksheet and delete the one with a given name using Aspose.Cells for .NET. | Provide a C# snippet that checks whether a pivot table exists at a specific index before calling RemoveAt to prevent exceptions. | Explain how to safely remove multiple pivot tables in a loop with PivotTables.RemoveAt, handling index shifts correctly.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This example creates a workbook, adds three pivot tables to the first worksheet, deletes the second table using sheet.PivotTables.RemoveAt(1), and saves the result as DeletedPivotTable.xlsx.
class DeletePivotTableDemo
{
    public static void Main()
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

        // Populate sample data for the pivot tables
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["B4"].PutValue(300);

        // Add three pivot tables to the worksheet
        sheet.PivotTables.Add("A1:B4", "D1",  "PivotTable1");
        sheet.PivotTables.Add("A1:B4", "D10", "PivotTable2");
        sheet.PivotTables.Add("A1:B4", "D20", "PivotTable3");

        // Remove the second pivot table (index 1) using RemoveAt
        sheet.PivotTables.RemoveAt(1);

        // Save the workbook to a file
        workbook.Save("DeletedPivotTable.xlsx");
        Console.WriteLine("Workbook saved as DeletedPivotTable.xlsx");
    }
}
