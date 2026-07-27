// Title: Aspose.Cells .NET – Add a Shape and Position It by Column‑Width Pixel Offset
// Description: C# sample that creates a workbook, computes the total pixel width of columns before a target column with Worksheet.Cells.GetColumnWidthPixel, adds a rectangle shape at column 0 using that pixel offset as the left position, prints expected vs. actual coordinates, and saves the file as ShapePositionDemo.xlsx.
// Keywords: Aspose.Cells | .NET | C# shape positioning | GetColumnWidthPixel | shape.Left property | absolute pixel offset | column width calculation | add rectangle shape | worksheet shapes | shape placement verification | Aspose.Cells example | GitHub Aspose.Cells shape
// Common Searches: Aspose.Cells position shape by column width | calculate pixel offset for shape placement .NET | set shape.Left using GetColumnWidthPixel | verify shape coordinates after adding Aspose.Cells | add rectangle to worksheet with absolute position
// Developer Intent: Place a shape at a precise horizontal location by converting a column index into pixel coordinates.
// Use Cases: Align a logo or banner with column D so it stays aligned when column widths change. | Generate a header graphic that starts at any chosen column by computing its pixel offset. | Automated test that confirms a shape's Left property matches the expected pixel calculation.
// AI Prompts: Write C# code that adds a circle shape and positions it at column index 5 using pixel offsets with Aspose.Cells. | Show how to retrieve and compare a shape's UpperLeftColumn and Left properties after placement in Aspose.Cells. | Explain how to adjust a shape's Top position based on row‑height pixels in a similar manner.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapePositionDemo
{
    // C# sample that creates a workbook, computes the total pixel width of columns before a target column with Worksheet.Cells.GetColumnWidthPixel, adds a rectangle shape at column 0 using that pixel offset as the left position, prints expected vs. actual coordinates, and saves the file as ShapePositionDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the target column index (zero‑based)
            int targetColumnIndex = 3; // e.g., column D

            // Calculate the absolute pixel offset from the left border
            // by summing the pixel widths of all columns before the target column
            int pixelOffset = 0;
            for (int col = 0; col < targetColumnIndex; col++)
            {
                pixelOffset += worksheet.Cells.GetColumnWidthPixel(col);
            }

            // Add a rectangle shape.
            // Place it at column 0 with a left offset equal to the calculated pixel offset.
            // Top row is set to 0 and top offset to 0 for simplicity.
            Shape shape = worksheet.Shapes.AddRectangle(
                topRow: 0,
                top: 0,
                leftColumn: 0,
                left: pixelOffset,
                height: 100,
                width: 200);

            // Test placement: verify that the shape's Left property matches the calculated offset
            // and that its UpperLeftColumn is 0 (as we placed it in column 0).
            Console.WriteLine("Expected Left offset (pixels): " + pixelOffset);
            Console.WriteLine("Actual Shape.Left: " + shape.Left);
            Console.WriteLine("Shape.UpperLeftColumn: " + shape.UpperLeftColumn);

            // Save the workbook
            workbook.Save("ShapePositionDemo.xlsx");
        }
    }
}
