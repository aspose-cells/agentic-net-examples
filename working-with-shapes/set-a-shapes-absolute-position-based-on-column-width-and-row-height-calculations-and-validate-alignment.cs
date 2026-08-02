// Title: Aspose.Cells for .NET – Position a Shape with Pixel Offsets from Column Widths & Row Heights and Validate Alignment
// Description: This C# example creates a workbook, sets custom column widths and row heights, calculates the cumulative pixel offset for a target cell, adds a rectangle shape, assigns its X/Y coordinates using those offsets, anchors the shape to the cell, sets Placement to MoveAndSize, validates the anchoring, aligns the shape's top‑right corner to another cell, and saves the file.
// Keywords: Aspose.Cells shape positioning | C# absolute shape coordinates | pixel offset column width | pixel offset row height | MoveAndSize placement | AlignTopRightCorner Aspose.Cells | shape anchoring Excel .NET | calculate shape X Y Aspose.Cells | Aspose.Cells AddRectangle example
// Common Searches: Aspose.Cells set shape X Y pixel offset | How to anchor a shape to a specific cell in Aspose.Cells .NET | Align shape top right corner to another cell Aspose.Cells | MoveAndSize placement for shapes Aspose.Cells | Get column width in pixels Aspose.Cells
// Developer Intent: Place a shape at an exact location based on column widths and row heights, anchor it to a cell, and ensure it moves/resizes with that cell.
// Use Cases: Insert a company logo that stays aligned with a header cell when column sizes change. | Create dynamic callout shapes whose top‑right corner points to a data cell for automated reports. | Ensure chart or image shapes follow the size and position of their anchor cells during worksheet edits.
// AI Prompts: Generate C# code with Aspose.Cells to position a shape at cell D7 by computing pixel offsets from column widths and row heights, then set Placement to MoveAndSize. | Write a method that aligns any shape's bottom‑left corner to a given cell while preserving its original dimensions in Aspose.Cells for .NET. | Explain how to retrieve column width and row height in pixels with Aspose.Cells and use them to calculate absolute X/Y coordinates for shape placement.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapePositionDemo
{
    // This C# example creates a workbook, sets custom column widths and row heights, calculates the cumulative pixel offset for a target cell, adds a rectangle shape, assigns its X/Y coordinates using those offsets, anchors the shape to the cell, sets Placement to MoveAndSize, validates the anchoring, aligns the shape's top‑right corner to another cell, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set custom column widths (in characters) and row heights (in points) for demonstration
            worksheet.Cells.SetColumnWidth(0, 15); // Column A
            worksheet.Cells.SetColumnWidth(1, 20); // Column B
            worksheet.Cells.SetColumnWidth(2, 25); // Column C
            worksheet.Cells.SetRowHeight(0, 30);   // Row 1
            worksheet.Cells.SetRowHeight(1, 40);   // Row 2
            worksheet.Cells.SetRowHeight(2, 50);   // Row 3
            worksheet.Cells.SetRowHeight(3, 60);   // Row 4
            worksheet.Cells.SetRowHeight(4, 70);   // Row 5

            // Target cell where the shape's upper‑left corner should be anchored
            int targetRow = 4;      // zero‑based index (Row 5 in Excel)
            int targetColumn = 2;   // zero‑based index (Column C in Excel)

            // Calculate the absolute pixel offset from the worksheet's left/top borders
            double leftPixel = 0;
            for (int c = 0; c < targetColumn; c++)
                leftPixel += worksheet.Cells.GetColumnWidthPixel(c);

            double topPixel = 0;
            for (int r = 0; r < targetRow; r++)
                topPixel += worksheet.Cells.GetRowHeightPixel(r);

            // Add a rectangle shape (initially at (0,0) with size 100x50 pixels)
            Shape shape = worksheet.Shapes.AddRectangle(0, 0, 100, 50, 0, 0);

            // Position the shape using absolute pixel coordinates
            shape.X = (int)leftPixel;   // horizontal offset from worksheet left border
            shape.Y = (int)topPixel;    // vertical offset from worksheet top border

            // Optionally, also set the cell anchoring properties for clarity
            shape.UpperLeftRow = targetRow;
            shape.UpperLeftColumn = targetColumn;
            shape.Left = 0;   // no additional offset inside the target column
            shape.Top = 0;    // no additional offset inside the target row

            // Set placement to MoveAndSize so the shape follows cell movements/resizing
            shape.Placement = PlacementType.MoveAndSize;

            // Validation: ensure the shape is anchored to the expected cell and placement
            if (shape.UpperLeftRow == targetRow &&
                shape.UpperLeftColumn == targetColumn &&
                shape.Placement == PlacementType.MoveAndSize)
            {
                Console.WriteLine("Shape positioned correctly at row {0}, column {1}.", targetRow + 1, targetColumn + 1);
            }
            else
            {
                Console.WriteLine("Shape positioning validation failed.");
            }

            // Demonstrate AlignTopRightCorner: align the shape's top‑right corner to another cell (row 2, column 5)
            shape.AlignTopRightCorner(1, 4); // zero‑based indices for row 2, column 5

            // Save the workbook
            workbook.Save("ShapeAbsolutePositionDemo.xlsx");
        }
    }
}
