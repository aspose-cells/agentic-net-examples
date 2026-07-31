// Title: Automatically add documentation comments to formula cells with Aspose.Cells (C#)
// Description: C# sample that creates or loads a workbook, inserts sample formulas, scans all used cells, and attaches a hidden comment to every formula cell. The comment records the formula, sets the author to "AutoDoc", and the workbook is saved as an annotated file.
// Keywords: Aspose.Cells add comment to formula | C# annotate Excel formulas | auto‑document formulas Aspose.Cells | iterate cells Aspose.Cells C# | programmatic Excel comments | hidden comment for audit Aspose | Excel formula documentation C#
// Common Searches: how to add a comment to each formula cell using Aspose.Cells C# | Aspose.Cells iterate over cells and insert comments | C# add hidden comments for Excel formulas | auto‑document Excel calculations with Aspose.Cells | programmatically annotate formula cells in .NET
// Developer Intent: Attach a comment to every cell that contains a formula to document its purpose.
// Use Cases: Create self‑documenting spreadsheets where each calculated cell shows its formula in a hidden note. | Add audit‑ready comments to financial models for traceability and reviewer insight. | Generate Excel templates that automatically annotate formula cells for downstream users or downstream automation.
// AI Prompts: Write C# code with Aspose.Cells that scans a worksheet and adds a comment containing the cell's formula and a custom description. | Modify the example so each comment includes a friendly explanation of the formula instead of just the formula string. | Create a reusable method that accepts a Workbook and a flag to add visible or hidden comments only to formula cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaComments
{
    // C# sample that creates or loads a workbook, inserts sample formulas, scans all used cells, and attaches a hidden comment to every formula cell. The comment records the formula, sets the author to "AutoDoc", and the workbook is saved as an annotated file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // Replace with new Workbook("input.xlsx") to load

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Example data: put some formulas for demonstration
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);
            worksheet.Cells["B1"].Formula = "=SUM(A1,A2)";          // simple sum
            worksheet.Cells["B2"].Formula = "=AVERAGE(A1:A2)";     // average
            worksheet.Cells["C1"].Formula = "=B1*2";              // uses result of B1

            // Iterate through all used cells
            Cells cells = worksheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the cell contains a formula
                    if (cell.IsFormula)
                    {
                        // Add a comment to the cell (if not already present)
                        int commentIndex = worksheet.Comments.Add(row, col);
                        Comment comment = worksheet.Comments[commentIndex];

                        // Set the comment text describing the formula purpose
                        // Here we simply record the formula itself; replace with custom description as needed
                        comment.Note = $"Formula: {cell.Formula}";
                        comment.Author = "AutoDoc";
                        comment.IsVisible = false; // hide by default
                    }
                }
            }

            // Save the workbook
            workbook.Save("FormulaCommentsDemo.xlsx");
        }
    }
}
