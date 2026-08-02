// Title: Remove a ListObject column and clear its formula with Aspose.Cells for .NET
// Description: Demonstrates how to delete a table (ListObject) column in a workbook, clear the column's Formula property, recalculate the sheet, and save the file so the removed column leaves no residual formulas in remaining rows.
// Keywords: Aspose.Cells delete ListObject column | clear table column formula .NET | remove Excel table column Aspose | recalculate workbook after column removal | C# Aspose.Cells ListObject column deletion
// Common Searches: how to delete a ListObject column in Aspose.Cells | clear formula before removing table column C# | Aspose.Cells remove column without leaving formulas | recalculate after deleting Excel table column Aspose | C# example delete ListObject column Aspose.Cells
// Developer Intent: Delete a ListObject column while ensuring its formula is cleared and does not affect other cells.
// Use Cases: Eliminate a temporary calculation column before exporting a report. | Clean up helper columns after performing intermediate data transformations. | Update a worksheet by removing a formula‑driven column and recalculating dependent values.
// AI Prompts: Write C# code using Aspose.Cells that clears a ListObject column's formula, removes the column, and recalculates the workbook. | Show a safe way to delete a table column in Aspose.Cells without leaving leftover formulas in other rows. | Explain the steps to remove a ListObject column and prevent its formula from propagating after deletion in a .NET workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to delete a table (ListObject) column in a workbook, clear the column's Formula property, recalculate the sheet, and save the file so the removed column leaves no residual formulas in remaining rows.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with a header row
        sheet.Cells["A1"].PutValue("Number");
        sheet.Cells["B1"].PutValue("Double");
        sheet.Cells["A2"].PutValue(5);
        sheet.Cells["A3"].PutValue(10);
        sheet.Cells["A4"].PutValue(15);

        // Add a ListObject (table) that includes the data range A1:B4
        int listIndex = sheet.ListObjects.Add("A1", "B4", true);
        ListObject listObj = sheet.ListObjects[listIndex];

        // Set a formula for the second column (index 1) of the table
        // This formula will be applied to all data rows in that column
        ListColumn doubleColumn = listObj.ListColumns[1];
        doubleColumn.Formula = "=A2*2";

        // Calculate formulas so the values are materialized
        workbook.CalculateFormula();

        // Before removal, the formula column exists
        Console.WriteLine("Before removal:");
        for (int row = 1; row <= sheet.Cells.MaxDataRow; row++)
        {
            Console.WriteLine($"Row {row + 1}: Number={sheet.Cells[row, 0].Value}, Double={sheet.Cells[row, 1].Value}");
        }

        // Clear the column formula to ensure it does not persist after deletion
        doubleColumn.Formula = null;

        // Remove the column from the ListObject
        // This also removes the underlying cells for that column
        listObj.ListColumns.RemoveAt(1);

        // Recalculate in case any dependent formulas exist
        workbook.CalculateFormula();

        // After removal, the formula column should no longer be present
        Console.WriteLine("\nAfter removal:");
        for (int row = 1; row <= sheet.Cells.MaxDataRow; row++)
        {
            // Column index 1 now refers to the original third column (if any)
            // Since we only had two columns, only column 0 remains
            Console.WriteLine($"Row {row + 1}: Number={sheet.Cells[row, 0].Value}");
        }

        // Save the modified workbook
        workbook.Save("RemovedListObjectColumn.xlsx");
    }
}
