using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ReadShadowColorDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 100);

                // Obtain the ShadowEffect object of the shape
                ShadowEffect shadowEffect = shape.ShadowEffect;

                // Create a CellsColor instance and assign a known color (e.g., Blue)
                CellsColor shadowColor = workbook.CreateCellsColor();
                shadowColor.Color = Color.Blue; // Set the shadow color

                // Apply the color to the shape's shadow
                shadowEffect.Color = shadowColor;

                // ----- Read the current shadow color -----
                // The Color property returns a CellsColor object
                CellsColor currentShadowColor = shadowEffect.Color;

                // Retrieve the System.Drawing.Color from the CellsColor
                Color rgbColor = currentShadowColor.Color;

                // Log the RGB components
                Console.WriteLine("Current Shadow Color (RGB):");
                Console.WriteLine($"Red   : {rgbColor.R}");
                Console.WriteLine($"Green : {rgbColor.G}");
                Console.WriteLine($"Blue  : {rgbColor.B}");

                // Save the workbook (lifecycle rule)
                string outputPath = "ReadShadowColorDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
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
            ReadShadowColorDemo.Run();
        }
    }
}