using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape (placeholder for a SmartArt shape)
            Shape shape = worksheet.Shapes.AddRectangle(0, 0, 200, 200, 0, 0);

            // Convert the shape to a GroupShape using GetResultOfSmartArt
            GroupShape result = shape.GetResultOfSmartArt();

            // Verify that the conversion result is a GroupShape
            if (result != null && result is GroupShape)
            {
                Console.WriteLine("Conversion result is a GroupShape.");
            }
            else
            {
                Console.WriteLine("Conversion result is not a GroupShape.");
            }

            // Save the workbook (lifecycle save)
            workbook.Save("SmartArtConversionResult.xlsx");
        }
    }
}