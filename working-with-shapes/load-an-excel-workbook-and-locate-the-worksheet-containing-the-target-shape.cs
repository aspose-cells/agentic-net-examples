using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class LocateShapeWorksheetDemo
    {
        public static void Run()
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Name of the shape to locate
                string targetShapeName = "MyRectangle";

                Worksheet targetWorksheet = null;

                // Search each worksheet for the shape
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    ShapeCollection shapes = sheet.Shapes;

                    // Attempt to retrieve the shape by name
                    Shape shape = null;
                    try
                    {
                        shape = shapes[targetShapeName];
                    }
                    catch
                    {
                        // Indexer throws if not found; ignore and continue
                    }

                    if (shape != null)
                    {
                        targetWorksheet = shape.Worksheet;
                        break; // Shape found
                    }
                }

                // Output the result
                if (targetWorksheet != null)
                {
                    Console.WriteLine($"Shape \"{targetShapeName}\" is located in worksheet: {targetWorksheet.Name}");
                }
                else
                {
                    Console.WriteLine($"Shape \"{targetShapeName}\" was not found in any worksheet.");
                }
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public static class Program
    {
        public static void Main(string[] args)
        {
            LocateShapeWorksheetDemo.Run();
        }
    }
}