// Title: Place a Shape Over a Merged Cell Range and Verify Its Position – Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, merges cells B2:D4, adds a rectangle shape, moves it to cover the merged area with MoveToRange, retrieves the merged range via GetMergedRange, compares the shape's UpperLeft/LowerRight coordinates with the range bounds, prints the alignment result, and saves the file.
// Keywords: Aspose.Cells shape over merged cells | MoveToRange C# | GetMergedRange verification | shape alignment Excel .NET | overlay shape merged range
// Common Searches: Aspose.Cells align shape with merged cells | C# MoveToRange merged range example | verify shape covers merged area Aspose.Cells | GetMergedRange usage for shapes
// Developer Intent: Position a shape so it exactly covers a merged cell block and programmatically confirm that the shape’s absolute coordinates match the merged range.
// Use Cases: Add a banner rectangle over a merged header in an automated report. | Validate that icons or images placed on a template align with merged cells before publishing. | Ensure consistent layout when programmatically overlaying shapes on merged regions in Excel workbooks.
// AI Prompts: Write C# code using Aspose.Cells to add a rectangle that covers merged cells B2:D4 and check its alignment. | Explain how MoveToRange and GetMergedRange work together to position shapes on merged ranges in Aspose.Cells for .NET. | Show best‑practice error handling when moving shapes onto merged cells with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

// This example creates a workbook, merges cells B2:D4, adds a rectangle shape, moves it to cover the merged area with MoveToRange, retrieves the merged range via GetMergedRange, compares the shape's UpperLeft/LowerRight coordinates with the range bounds, prints the alignment result, and saves the file.
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

            // Define the merged range (B2:D4) – zero‑based indices
            int firstRow = 1;        // B2 row index
            int firstColumn = 1;     // B2 column index
            int totalRows = 3;       // rows 2,3,4
            int totalColumns = 3;    // columns B,C,D

            // Merge the cells
            cells.Merge(firstRow, firstColumn, totalRows, totalColumns);

            // Add a rectangle shape (initial position does not matter)
            Shape shape = worksheet.Shapes.AddRectangle(0, 0, 50, 50, 0, 0);

            // Move the shape so that it exactly covers the merged range
            shape.MoveToRange(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1);

            // Retrieve the merged range via the top‑left cell
            Cell topLeftCell = cells[firstRow, firstColumn];
            AsposeRange mergedRange = topLeftCell.GetMergedRange();

            // Verify that the shape's absolute position matches the merged area
            bool aligns =
                shape.UpperLeftRow == mergedRange.FirstRow &&
                shape.UpperLeftColumn == mergedRange.FirstColumn &&
                shape.LowerRightRow == mergedRange.FirstRow + mergedRange.RowCount - 1 &&
                shape.LowerRightColumn == mergedRange.FirstColumn + mergedRange.ColumnCount - 1;

            Console.WriteLine("Shape aligns with merged area: " + aligns);

            // Save the workbook
            workbook.Save("ShapeOverMergedCell.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
