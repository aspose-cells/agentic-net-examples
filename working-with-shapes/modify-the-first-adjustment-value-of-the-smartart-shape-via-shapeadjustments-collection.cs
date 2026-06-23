using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ModifySmartArtAdjustmentDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a chevron auto shape (supports adjustment guides)
                Shape shape = worksheet.Shapes.AddAutoShape(AutoShapeType.Chevron, 10, 10, 0, 0, 200, 100);

                // Access the geometry which holds adjustment values
                Geometry geometry = shape.Geometry;

                if (geometry.ShapeAdjustValues.Count > 0)
                {
                    // Modify the first adjustment value
                    geometry.ShapeAdjustValues[0].Value = 0.4;
                    Console.WriteLine("First adjustment value modified to 0.4");
                }
                else
                {
                    // Add a new adjustment guide and set its value
                    int guideIndex = geometry.ShapeAdjustValues.Add("adj1", 0.4);
                    Console.WriteLine($"Added adjustment guide at index {guideIndex} with value 0.4");
                }

                // Save the workbook
                string outputPath = "ModifySmartArtAdjustmentDemo.xlsx";
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
            ModifySmartArtAdjustmentDemo.Run();
        }
    }
}