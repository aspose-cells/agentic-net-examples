using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ApplyBlueGlowToAllShapes
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample shapes to demonstrate the effect
            // Rectangle
            worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 0, 0, 0, 100, 100);
            // Oval (ellipse)
            worksheet.Shapes.AddShape(MsoDrawingType.Oval, 2, 0, 0, 0, 120, 80);
            // Rounded rectangle (auto shape)
            worksheet.Shapes.AddAutoShape(AutoShapeType.RoundedRectangle, 3, 0, 0, 0, 150, 70);

            // Apply a blue glow of size 8 points to every shape in the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Access the glow effect of the shape
                GlowEffect glow = shape.Glow;

                // Set the glow radius (size) to 8 points
                glow.Size = 8;

                // Set the glow color to blue
                CellsColor glowColor = workbook.CreateCellsColor();
                glowColor.Color = Color.Blue;
                glow.Color = glowColor;
            }

            // Define output file path
            string outputPath = "AllShapesBlueGlow.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the applied glow effects
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}