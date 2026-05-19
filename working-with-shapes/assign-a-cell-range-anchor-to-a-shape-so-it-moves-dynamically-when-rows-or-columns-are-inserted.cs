using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ShapeAnchorDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet (initial position is arbitrary)
            Shape shape = worksheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);

            // Set the shape's placement so it moves and resizes with the cells
            shape.Placement = PlacementType.MoveAndSize;

            // Anchor the shape to a specific cell range (e.g., B2:D5)
            // This makes the shape adjust its position when rows/columns are inserted
            shape.MoveToRange(1, 1, 4, 3); // rows and columns are zero‑based indices

            // Define output file path
            string outputPath = "ShapeAnchorDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ShapeAnchorDemo.Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}