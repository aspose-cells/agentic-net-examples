using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeOverMergedCellDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the merged range (e.g., B2:D4 -> rows 1‑3, columns 1‑3)
            int firstRow = 1;      // zero‑based index for row 2
            int firstColumn = 1;   // zero‑based index for column B
            int totalRows = 3;     // rows 2,3,4
            int totalColumns = 3;  // columns B,C,D

            // Merge the cells
            cells.Merge(firstRow, firstColumn, totalRows, totalColumns);

            // Add a rectangle shape (initial position does not matter)
            Shape shape = worksheet.Shapes.AddRectangle(0, 0, 50, 50, 0, 0);

            // Move the shape so that it covers the merged range
            shape.MoveToRange(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1);

            // Retrieve the merged range via the top‑left cell
            Cell topLeftCell = cells[firstRow, firstColumn];
            Aspose.Cells.Range mergedRange = topLeftCell.GetMergedRange();

            // Verify that the shape's absolute position matches the merged area
            bool aligns =
                shape.UpperLeftRow == mergedRange.FirstRow &&
                shape.UpperLeftColumn == mergedRange.FirstColumn &&
                shape.LowerRightRow == mergedRange.FirstRow + mergedRange.RowCount - 1 &&
                shape.LowerRightColumn == mergedRange.FirstColumn + mergedRange.ColumnCount - 1;

            Console.WriteLine("Shape aligns with merged area: " + aligns);

            // Save the workbook (ensure the directory exists)
            string outputPath = "ShapeOverMergedCell.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}