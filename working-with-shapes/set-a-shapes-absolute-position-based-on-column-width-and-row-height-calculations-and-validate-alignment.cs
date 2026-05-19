using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapePositionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set custom column widths (in pixels) for demonstration
            // Column 0: 80px, Column 1: 120px, Column 2: 100px
            worksheet.Cells.SetColumnWidthPixel(0, 80);
            worksheet.Cells.SetColumnWidthPixel(1, 120);
            worksheet.Cells.SetColumnWidthPixel(2, 100);

            // Set custom row heights (in pixels) for demonstration
            // Row 0: 30px, Row 1: 45px, Row 2: 60px
            worksheet.Cells.SetRowHeightPixel(0, 30);
            worksheet.Cells.SetRowHeightPixel(1, 45);
            worksheet.Cells.SetRowHeightPixel(2, 60);

            // Desired cell location for the shape's upper‑left corner
            int targetRow = 1;      // second row (zero‑based)
            int targetColumn = 2;   // third column (zero‑based)

            // Calculate the absolute X offset (pixels) from the worksheet's left border
            int absoluteX = 0;
            for (int col = 0; col < targetColumn; col++)
            {
                absoluteX += worksheet.Cells.GetColumnWidthPixel(col);
            }

            // Calculate the absolute Y offset (pixels) from the worksheet's top border
            int absoluteY = 0;
            for (int row = 0; row < targetRow; row++)
            {
                absoluteY += worksheet.Cells.GetRowHeightPixel(row);
            }

            // Add a rectangle shape (initial position values are placeholders)
            Shape shape = worksheet.Shapes.AddRectangle(0, 0, 0, 150, 80, 0);

            // Set the shape's cell anchor (upper‑left corner)
            shape.UpperLeftRow = targetRow;
            shape.UpperLeftColumn = targetColumn;

            // Apply the calculated pixel offsets within the anchored cell
            shape.Top = absoluteY - worksheet.Cells.GetRowHeightPixel(targetRow) + 5;   // 5px offset inside the cell
            shape.Left = absoluteX - worksheet.Cells.GetColumnWidthPixel(targetColumn) + 10; // 10px offset inside the cell

            // Validate alignment: ensure the shape's placement type moves and sizes with cells
            shape.Placement = PlacementType.MoveAndSize;

            // Simple validation output
            Console.WriteLine($"Shape anchored at row {shape.UpperLeftRow}, column {shape.UpperLeftColumn}");
            Console.WriteLine($"Shape pixel offset - Top: {shape.Top}, Left: {shape.Left}");
            Console.WriteLine($"Placement type: {shape.Placement}");

            // Save the workbook
            workbook.Save("ShapeAbsolutePositionDemo.xlsx");
        }
    }
}