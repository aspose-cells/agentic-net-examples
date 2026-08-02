// Title: Remove a specific PivotTable using PivotTables.RemoveAt in C# (Aspose.Cells)
// Description: This C# sample builds a workbook, inserts three pivot tables, then deletes the pivot table at zero‑based index 1 with the PivotTables.RemoveAt method and saves the result as DeletedPivotTable.xlsx, demonstrating how to programmatically eliminate a pivot table by its position using Aspose.Cells for .NET.
// Keywords: Aspose.Cells remove pivot table C# | PivotTables.RemoveAt example | delete pivot table by index Aspose.Cells | C# Aspose.Cells PivotTable removal | remove specific pivot table .NET | Aspose.Cells PivotTable Delete | PivotTable.RemoveAt usage | Aspose.Cells workbook cleanup
// Common Searches: how to delete a pivot table with Aspose.Cells C# | PivotTables.RemoveAt usage example | remove second pivot table Aspose.Cells | Aspose.Cells delete pivot table by index | C# code to remove a specific PivotTable
// Developer Intent: Programmatically delete a pivot table from a worksheet by its zero‑based index using Aspose.Cells for .NET.
// Use Cases: Clean up a report after generating multiple pivot tables, keeping only the needed ones. | Automate workbook maintenance by removing outdated pivot tables before distribution. | Replace an old pivot table with a refreshed version without recreating the entire sheet.
// AI Prompts: Generate C# code that removes the third PivotTable in a worksheet using Aspose.Cells. | Show how to verify the count of PivotTables before calling RemoveAt to prevent out‑of‑range errors. | Provide an example that deletes a PivotTable and then renames the remaining tables in the same workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This C# sample builds a workbook, inserts three pivot tables, then deletes the pivot table at zero‑based index 1 with the PivotTables.RemoveAt method and saves the result as DeletedPivotTable.xlsx, demonstrating how to programmatically eliminate a pivot table by its position using Aspose.Cells for .NET.
class DeletePivotTableDemo
{
    static void Main()
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

        // Delete the second pivot table (index 1) using RemoveAt
        sheet.PivotTables.RemoveAt(1);

        // Save the workbook to verify the pivot table removal
        workbook.Save("DeletedPivotTable.xlsx");
    }
}
