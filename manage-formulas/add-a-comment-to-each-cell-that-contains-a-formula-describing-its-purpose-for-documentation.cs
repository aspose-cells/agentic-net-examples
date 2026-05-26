using System;
using Aspose.Cells;

namespace AsposeCellsFormulaComments
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all used cells in the worksheet
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
                        // Get the cell name (e.g., "A1")
                        string cellName = CellsHelper.CellIndexToName(row, col);

                        // Add a comment to the cell (or retrieve existing one)
                        int commentIndex = worksheet.Comments.Add(cellName);
                        Comment comment = worksheet.Comments[commentIndex];

                        // Set the comment text describing the formula purpose
                        // Here we simply store the formula itself; customize as needed
                        comment.Note = $"Formula: {cell.Formula}";
                    }
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}