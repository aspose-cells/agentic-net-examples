// Title: Add comments to every formula cell in an Aspose.Cells workbook using C#
// Description: Creates a new workbook, populates cells with values and formulas, then scans the used range. For each cell where IsFormula is true, a comment is added via worksheet.Comments.Add, containing the formula text, a brief purpose note, and an author tag. The workbook is saved with the documentation embedded.
// Keywords: Aspose.Cells C# add comment | document Excel formulas programmatically | iterate used cells Aspose.Cells | auto‑generate cell comments | formula annotation .NET | Excel workbook documentation | C# comment author Aspose
// Common Searches: how to add comments to formula cells with Aspose.Cells .NET | C# iterate over used range and annotate formulas | Aspose.Cells add note to calculated cells | programmatically document Excel formulas C# | add author to Excel cell comment Aspose
// Developer Intent: Programmatically attach a descriptive comment to each cell that contains a formula, providing documentation and author attribution.
// Use Cases: Create an audit trail for financial models by embedding explanatory notes directly in calculated cells. | Enhance readability of complex spreadsheets for end users by automatically generating formula descriptions. | Standardize documentation across exported reports, ensuring every calculated field carries an author‑identified comment.
// AI Prompts: Generate C# code with Aspose.Cells that adds a custom comment to each formula cell, including the formula string and a concise explanation. | Adapt the example to assign different comment authors based on worksheet names while preserving the formula note. | Design a reusable method that accepts a Workbook and a comment template, then applies comments to all formula cells in the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaComments
{
    // Creates a new workbook, populates cells with values and formulas, then scans the used range. For each cell where IsFormula is true, a comment is added via worksheet.Comments.Add, containing the formula text, a brief purpose note, and an author tag. The workbook is saved with the documentation embedded.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data with formulas for demonstration
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["B1"].PutValue(5);
            cells["B2"].PutValue(15);
            cells["C1"].Formula = "=SUM(A1:A2)";          // Formula cell
            cells["C2"].Formula = "=AVERAGE(B1,B2)";     // Formula cell
            cells["D1"].Formula = "=C1*B1";              // Formula cell

            // Iterate through all used cells to find formulas
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.IsFormula) // Check if the cell contains a formula
                    {
                        // Add a comment to the cell using the Add(row, column) method
                        int commentIndex = worksheet.Comments.Add(row, col);
                        Comment comment = worksheet.Comments[commentIndex];

                        // Describe the purpose of the formula (customize as needed)
                        comment.Note = $"Formula: {cell.Formula} – this cell computes the required value.";
                        comment.Author = "AutoDoc";
                    }
                }
            }

            // Save the workbook with the added comments
            workbook.Save("FormulaCommentsDemo.xlsx");
        }
    }
}
