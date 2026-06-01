using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeInfoDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in‑memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a few sample shapes so the demo has something to list
            // Rectangle at row 2, column 2, size 100x50 pixels
            sheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 50);
            // Oval at row 5, column 5, size 80x80 pixels
            sheet.Shapes.AddOval(5, 0, 5, 0, 80, 80);
            // TextBox at row 8, column 3, size 150x40 pixels
            sheet.Shapes.AddTextBox(8, 0, 3, 0, 150, 40);

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Shape name (may be empty if not set)
                string name = shape.Name;

                // Drawing type (Rectangle, Oval, TextBox, etc.)
                MsoDrawingType drawingType = shape.MsoDrawingType;

                // Absolute position and size in pixels
                int left   = shape.Left;   // X‑coordinate from the left edge of the sheet
                int top    = shape.Top;    // Y‑coordinate from the top edge of the sheet
                int width  = shape.Width;  // Width in pixels
                int height = shape.Height; // Height in pixels

                Console.WriteLine($"Name: {name}");
                Console.WriteLine($"Type: {drawingType}");
                Console.WriteLine($"Position: Left={left}px, Top={top}px");
                Console.WriteLine($"Size: Width={width}px, Height={height}px");
                Console.WriteLine(new string('-', 40));
            }

            // Save the workbook to verify that shapes are persisted
            workbook.Save("ShapesInfoDemo.xlsx");
        }
    }
}