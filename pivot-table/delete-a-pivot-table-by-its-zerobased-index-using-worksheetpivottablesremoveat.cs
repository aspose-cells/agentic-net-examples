// Title: Remove a Pivot Table by Index with Worksheet.PivotTables.RemoveAt (Aspose.Cells for .NET C#)
// Description: This C# example shows how to create a workbook, add three pivot tables, delete the second pivot table using Worksheet.PivotTables.RemoveAt(1), verify the remaining count, and save the file as PivotTableRemoved.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Worksheet.PivotTables.RemoveAt | delete pivot table | remove pivot table by index | pivot table API | Aspose.Cells pivot table | remove specific pivot table | Aspose.Cells example
// Common Searches: Aspose.Cells remove pivot table by index | Worksheet.PivotTables.RemoveAt C# example | delete specific pivot table Aspose.Cells .NET | how to remove second pivot table Aspose.Cells | remove pivot table programmatically Aspose.Cells
// Developer Intent: Delete a pivot table from a worksheet using its zero‑based index.
// Use Cases: Eliminate unwanted pivot tables generated during dynamic reporting | Reduce workbook size by removing temporary pivot tables before export | Replace an outdated pivot table by deleting it at a known position and inserting a new one | Clean up pivot tables in automated spreadsheet processing pipelines
// AI Prompts: Show code to check the pivot table count before and after calling RemoveAt. | Provide a loop that removes all pivot tables from a worksheet. | Explain how to locate a pivot table's index by name and delete it with RemoveAt. | Demonstrate error handling when the specified index is out of range. | Give an example of removing pivot tables from multiple worksheets in a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDeleteDemo
{
    // This C# example shows how to create a workbook, add three pivot tables, delete the second pivot table using Worksheet.PivotTables.RemoveAt(1), verify the remaining count, and save the file as PivotTableRemoved.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for pivot tables
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(300);

            // Add three pivot tables to the worksheet
            sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            sheet.PivotTables.Add("A1:B4", "D10", "PivotTable2");
            sheet.PivotTables.Add("A1:B4", "D20", "PivotTable3");

            // Remove the pivot table at zero‑based index 1 (the second pivot table)
            sheet.PivotTables.RemoveAt(1);

            // Optional: verify the remaining count
            Console.WriteLine("Remaining Pivot Tables Count: " + sheet.PivotTables.Count);

            // Save the workbook
            workbook.Save("PivotTableRemoved.xlsx");
        }
    }
}
