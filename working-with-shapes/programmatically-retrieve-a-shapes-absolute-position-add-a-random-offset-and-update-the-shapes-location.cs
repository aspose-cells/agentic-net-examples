using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapePositionAdjustment
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape; otherwise add a sample rectangle
            if (worksheet.Shapes.Count == 0)
            {
                // AddRectangle parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
                worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 50);
            }

            // Retrieve the first shape
            Shape shape = worksheet.Shapes[0];

            // Get the shape's current absolute position (pixels from worksheet borders)
            int currentX = shape.X; // Horizontal offset from left border
            int currentY = shape.Y; // Vertical offset from top border

            // Generate random offsets (e.g., between -20 and +20 pixels)
            Random rnd = new Random();
            int offsetX = rnd.Next(-20, 21);
            int offsetY = rnd.Next(-20, 21);

            // Apply the random offsets to the shape's position
            shape.X = currentX + offsetX;
            shape.Y = currentY + offsetY;

            // Optional: output the before/after positions to console for verification
            Console.WriteLine($"Original Position: X={currentX}, Y={currentY}");
            Console.WriteLine($"Offset Applied:   X={offsetX}, Y={offsetY}");
            Console.WriteLine($"New Position:     X={shape.X}, Y={shape.Y}");

            // Save the workbook with the updated shape position
            workbook.Save("output.xlsx");
        }
    }
}