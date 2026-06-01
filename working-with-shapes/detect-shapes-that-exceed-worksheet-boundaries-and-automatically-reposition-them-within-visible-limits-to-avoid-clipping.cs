using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeBoundaryAdjustment
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Work with the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the visible limits of the worksheet.
            // Here we use the maximum used row and column as the boundary.
            int maxRow = sheet.Cells.MaxDataRow;      // zero‑based index of the last used row
            int maxColumn = sheet.Cells.MaxDataColumn; // zero‑based index of the last used column

            // Iterate through all shapes on the worksheet
            for (int i = 0; i < sheet.Shapes.Count; i++)
            {
                Shape shape = sheet.Shapes[i];

                // Get the current top‑left cell of the shape
                int shapeRow = shape.UpperLeftRow;
                int shapeColumn = shape.UpperLeftColumn;

                bool needsAdjustment = false;

                // If the shape starts beyond the last used row, move it up
                if (shapeRow > maxRow)
                {
                    shapeRow = maxRow;
                    needsAdjustment = true;
                }

                // If the shape starts beyond the last used column, move it left
                if (shapeColumn > maxColumn)
                {
                    shapeColumn = maxColumn;
                    needsAdjustment = true;
                }

                // Ensure the shape does not start before the first row/column
                if (shapeRow < 0)
                {
                    shapeRow = 0;
                    needsAdjustment = true;
                }

                if (shapeColumn < 0)
                {
                    shapeColumn = 0;
                    needsAdjustment = true;
                }

                // Apply the new position if any adjustment was required
                if (needsAdjustment)
                {
                    // MoveToRange moves the shape to a new range.
                    // Here we keep the shape size unchanged by using the same start and end cell.
                    shape.MoveToRange(shapeRow, shapeColumn, shapeRow, shapeColumn);
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}