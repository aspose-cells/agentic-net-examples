using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtValidation
{
    public class ValidateIsSmartArt
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a regular rectangle shape (non‑SmartArt)
                // Parameters: upperLeftRow, upperLeftColumn, upperLeftPixelRow, upperLeftPixelColumn, height, width
                Shape rectangle = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 100);

                // Verify that IsSmartArt is false for this regular shape
                bool isSmartArt = rectangle.IsSmartArt;
                Console.WriteLine($"Rectangle IsSmartArt: {isSmartArt}"); // Expected output: False

                // Save the workbook (optional, demonstrates lifecycle usage)
                string outputPath = "IsSmartArtValidation.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateIsSmartArt.Run();
        }
    }
}