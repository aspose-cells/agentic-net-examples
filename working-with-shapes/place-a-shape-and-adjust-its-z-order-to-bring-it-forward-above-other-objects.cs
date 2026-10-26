// Title: Add a rectangle shape and set its Z‑order to bring it to the front in an Excel file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add rectangle shape and change ZOrderPosition | how to set shape Z-order in Excel with Aspose.Cells .NET | bring a worksheet shape to front using Aspose.Cells API | C# code to layer shapes in an Excel workbook with Aspose.Cells | Aspose.Cells shape layering example C#
// Tags: add rectangle shape Aspose.Cells | set shape ZOrderPosition Aspose.Cells | worksheet shape front ordering | C# Excel shape layering | Aspose.Cells shape Z-order example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape to the first worksheet, assigns its ZOrderPosition to the highest value to bring it forward, and saves the workbook as ShapeZOrder.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape (row index 1, column index 0, offsetX 0, offsetY 0, width 100, height 50)
            // The AddShape method returns the created Shape object
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,    // upper left row
                0,    // upper left column
                0,    // top offset (in pixels)
                0,    // left offset (in pixels)
                50,   // height (in pixels)
                100); // width (in pixels)

            // Bring the shape to the front by setting a high Z‑order position
            shape.ZOrderPosition = sheet.Shapes.Count; // highest Z‑order

            // Define output file path
            string outputPath = "ShapeZOrder.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
