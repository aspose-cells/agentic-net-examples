using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsAdjustmentComparison
{
    public class Program
    {
        public static void Main()
        {
            // ---------- Create a new workbook and add a shape ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add an auto shape (Chevron) to the worksheet
            Shape shape = sheet.Shapes.AddAutoShape(AutoShapeType.Chevron, 10, 10, 0, 0, 200, 100);
            Geometry geometry = shape.Geometry;

            // Get the number of adjustment values before any conversion
            int adjustCountBefore = geometry.ShapeAdjustValues.Count;
            Console.WriteLine("Adjustment values count before conversion: " + adjustCountBefore);

            // ---------- Save the workbook ----------
            string filePath = "AdjustmentComparison.xlsx";
            workbook.Save(filePath);

            // ---------- Load the workbook back ----------
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // Retrieve the same shape (first shape in the collection)
            Shape loadedShape = loadedSheet.Shapes[0];
            Geometry loadedGeometry = loadedShape.Geometry;

            // Get the number of adjustment values after loading (conversion)
            int adjustCountAfter = loadedGeometry.ShapeAdjustValues.Count;
            Console.WriteLine("Adjustment values count after conversion: " + adjustCountAfter);

            // ---------- Compare the counts ----------
            if (adjustCountBefore == adjustCountAfter)
            {
                Console.WriteLine("Data integrity verified: counts are equal.");
            }
            else
            {
                Console.WriteLine("Data integrity issue: counts differ.");
            }
        }
    }
}