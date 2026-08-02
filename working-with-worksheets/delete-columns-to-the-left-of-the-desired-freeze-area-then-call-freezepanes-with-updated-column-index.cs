// Title: Delete Columns and Adjust FreezePanes Index with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, populate cells, delete leading columns, recalculate the freeze column index, apply FreezePanes at a specific row and column, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# FreezePanes | DeleteColumns Aspose.Cells | adjust freeze column after deletion | Excel freeze panes after column removal | Aspose.Cells workbook manipulation | C# delete columns then freeze panes
// Common Searches: Aspose.Cells delete columns then freeze panes C# | How to recalculate FreezePanes column index after deleting columns | C# example for FreezePanes with updated column index | Aspose.Cells remove leading columns before freezing rows | FreezePanes after DeleteColumns Aspose.Cells
// Developer Intent: Remove specific columns from a worksheet and set FreezePanes using the corrected row and column indices.
// Use Cases: Strip unwanted left‑most columns from a generated report while keeping header rows frozen. | Programmatically modify a template (e.g., delete placeholder columns) before applying FreezePanes for better user navigation. | Maintain correct formula references after column deletion and then freeze the view to keep key rows/columns visible.
// AI Prompts: Generate C# code with Aspose.Cells that deletes the first N columns, updates the freeze column index, and calls FreezePanes at row 5, column 5. | Explain the math for adjusting the FreezePanes column parameter after removing columns, including impact on formulas, in Aspose.Cells for .NET. | Show how to verify that the freeze area is correct after deleting columns and saving the workbook with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, populate cells, delete leading columns, recalculate the freeze column index, apply FreezePanes at a specific row and column, and save the file using Aspose.Cells for .NET.
    public class FreezePanesAfterDeletingColumnsDemo
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
            Cells cells = sheet.Cells;

            // Populate some sample data in columns A to E
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Desired freeze area: row 4 (zero‑based index) and column 4 (i.e., column "E")
            int desiredFreezeRow = 4;      // 5th row in Excel (index 4)
            int desiredFreezeColumn = 4;   // 5th column in Excel (index 4)

            // Suppose we need to delete the first two columns (A and B) before freezing
            int columnsToDelete = 2;
            int firstColumnToDelete = 0; // zero‑based index of column A

            // Delete the columns and update references in formulas
            cells.DeleteColumns(firstColumnToDelete, columnsToDelete, true);

            // After deletion, the original column index shifts left by the number of deleted columns
            int updatedFreezeColumn = desiredFreezeColumn - columnsToDelete;

            // Freeze panes using the updated column index
            // frozenRows and frozenColumns should match the freeze position
            sheet.FreezePanes(desiredFreezeRow, updatedFreezeColumn, desiredFreezeRow, updatedFreezeColumn);

            // Save the workbook
            string outputPath = "FreezePanesAfterDeletingColumns.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
