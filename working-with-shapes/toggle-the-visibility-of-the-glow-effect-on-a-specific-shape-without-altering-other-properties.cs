using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ToggleGlowVisibility
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, top, left, height, width
                Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);

                // Configure an initial glow effect (visible)
                shape.Glow.Size = 8;                 // radius in points
                shape.Glow.Transparency = 0.4;       // 40% transparent
                shape.Glow.Color = workbook.CreateCellsColor();
                shape.Glow.Color.Color = Color.Yellow;

                // Store the current glow size to restore later when toggling back on
                double originalSize = shape.Glow.Size;

                // Toggle logic: hide if visible, otherwise restore
                if (shape.Glow.Size > 0)
                {
                    shape.Glow.Size = 0; // hide glow
                }
                else
                {
                    shape.Glow.Size = originalSize; // show glow
                }

                // Save the workbook
                string outputPath = "ToggleGlowVisibility.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ToggleGlowVisibility.Run();
        }
    }
}