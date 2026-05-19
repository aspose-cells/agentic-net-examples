using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    public class AdjustZOrder
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add two rectangle shapes
                Shape shapeA = worksheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);   // Existing shape
                Shape shapeB = worksheet.Shapes.AddRectangle(20, 20, 100, 100, 0, 0); // Shape to be moved

                // Set initial Z-order positions (optional, defaults are sequential)
                shapeA.ZOrderPosition = 0; // bottom
                shapeB.ZOrderPosition = 1; // above shapeA

                // Adjust shapeB to be just above shapeA
                shapeB.ZOrderPosition = shapeA.ZOrderPosition + 1;

                // Save the workbook
                string outputPath = "AdjustedZOrder.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error adjusting Z-order: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            AdjustZOrder.Run();
        }
    }
}