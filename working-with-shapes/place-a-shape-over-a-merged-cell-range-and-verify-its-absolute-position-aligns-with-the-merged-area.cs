// Title: C# – Place and Verify a Shape Over a Merged Cell Range with Aspose.Cells
// Description: Creates a workbook, merges a block of cells, adds a rectangle shape, sets PlacementType.MoveAndSize, moves the shape to the merged area with MoveToRange, retrieves the merged range via GetMergedRange, compares UpperLeftRow/Column and LowerRightRow/Column to confirm exact alignment, and saves the file.
// Keywords: Aspose.Cells shape merged cells | MoveToRange C# example | GetMergedRange verification | PlacementType.MoveAndSize | shape alignment Aspose.Cells | C# Excel shape positioning
// Common Searches: Aspose.Cells place shape over merged cells | verify shape coordinates match merged range .NET | MoveToRange merged range Aspose.Cells C# | how to align rectangle with merged cells Aspose | shape placement type move and size example
// Developer Intent: Position a shape precisely over a merged cell block and programmatically confirm that its boundaries match the merged area.
// Use Cases: Add a header banner that spans a merged title row and stays aligned during column resizing. | Create an interactive button that covers a merged table section and validate its location before attaching a macro. | Insert a chart placeholder over a merged data range and ensure correct placement before rendering the chart.
// AI Prompts: Write C# code using Aspose.Cells to add a rectangle shape over a merged range and verify its UpperLeftRow/Column and LowerRightRow/Column match the merged area. | Explain how MoveToRange and GetMergedRange work together to align shapes with merged cells in Aspose.Cells for .NET. | Provide robust error‑handling patterns for positioning shapes over merged cells and checking alignment in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsShapeOverMergedCell
{
    // Creates a workbook, merges a block of cells, adds a rectangle shape, sets PlacementType.MoveAndSize, moves the shape to the merged area with MoveToRange, retrieves the merged range via GetMergedRange, compares UpperLeftRow/Column and LowerRightRow/Column to confirm exact alignment, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Define the merged range (starting at row 2, column 1, spanning 3 rows and 4 columns)
                int startRow = 2;      // zero‑based index (row 3 in Excel)
                int startColumn = 1;   // zero‑based index (column B)
                int totalRows = 3;
                int totalColumns = 4;

                // Merge the cells
                cells.Merge(startRow, startColumn, totalRows, totalColumns);

                // Add a rectangle shape (initial position will be overridden)
                Shape shape = worksheet.Shapes.AddRectangle(0, 0, 50, 50, 0, 0);

                // Ensure the shape moves and resizes with the cells
                shape.Placement = PlacementType.MoveAndSize;

                // Move the shape to exactly cover the merged range
                int endRow = startRow + totalRows - 1;
                int endColumn = startColumn + totalColumns - 1;
                shape.MoveToRange(startRow, startColumn, endRow, endColumn);

                // Retrieve the merged range information from the top‑left cell
                Cell topLeftCell = cells[startRow, startColumn];
                AsposeRange mergedRange = topLeftCell.GetMergedRange();

                // Calculate expected boundaries
                int expectedTopRow = mergedRange.FirstRow;
                int expectedLeftColumn = mergedRange.FirstColumn;
                int expectedBottomRow = mergedRange.FirstRow + mergedRange.RowCount - 1;
                int expectedRightColumn = mergedRange.FirstColumn + mergedRange.ColumnCount - 1;

                // Verify shape's absolute position matches the merged area
                bool isAligned = shape.UpperLeftRow == expectedTopRow &&
                                 shape.UpperLeftColumn == expectedLeftColumn &&
                                 shape.LowerRightRow == expectedBottomRow &&
                                 shape.LowerRightColumn == expectedRightColumn;

                Console.WriteLine(isAligned
                    ? "Shape is correctly aligned with the merged cell range."
                    : "Shape alignment does not match the merged cell range.");

                // Save the workbook
                workbook.Save("ShapeOverMergedCell.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
