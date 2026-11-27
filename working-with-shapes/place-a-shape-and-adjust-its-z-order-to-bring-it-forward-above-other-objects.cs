// Title: How to change the Z‑order of a rectangle shape in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# move shape forward in a worksheet | Change Z order of a rectangle with Aspose.Cells Drawing API | Bring a shape to front after sending it to back using ToFrontOrBack | Adjust shape layering in an Excel file with Aspose.Cells .NET | Use ToFrontOrBack to reorder shapes in a workbook
// Tags: Aspose.Cells shape Z-order | C# ToFrontOrBack method | Aspose.Cells rectangle layering | Excel workbook shape ordering .NET | Aspose.Cells Drawing API shape manipulation

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ZOrderDemo
{
    // The example creates a workbook, adds two rectangle shapes, uses the ToFrontOrBack method to send the foreground shape back and then bring it forward, and saves the file as ZOrderDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a first rectangle shape (will be the background shape)
                // Parameters: upper left row, upper left column, top offset, left offset, height, width
                Shape backgroundShape = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 150, 150);
                backgroundShape.Name = "BackgroundRect";
                backgroundShape.Fill.FillType = FillType.Solid;
                backgroundShape.Fill.SolidFill.Color = System.Drawing.Color.LightGray;
                backgroundShape.Line.Weight = 1;
                backgroundShape.Line.DashStyle = MsoLineDashStyle.Solid;

                // Add a second rectangle shape that will be placed on top of the first one
                Shape foregroundShape = worksheet.Shapes.AddRectangle(3, 0, 3, 0, 150, 150);
                foregroundShape.Name = "ForegroundRect";
                foregroundShape.Fill.FillType = FillType.Solid;
                foregroundShape.Fill.SolidFill.Color = System.Drawing.Color.CornflowerBlue;
                foregroundShape.Line.Weight = 2;
                foregroundShape.Line.DashStyle = MsoLineDashStyle.Solid;

                // At this point both shapes have the same Z‑order (the second added shape is already on top).
                // To demonstrate Z‑order manipulation, we will send the foreground shape to the back
                // and then bring it forward again.

                // Send the foreground shape 1 position back (negative orders)
                foregroundShape.ToFrontOrBack(-1);
                Console.WriteLine("Foreground shape sent to back.");

                // Bring the foreground shape 2 positions forward (positive orders)
                // Since there are only two shapes, this will place it at the very front.
                foregroundShape.ToFrontOrBack(2);
                Console.WriteLine("Foreground shape brought to front.");

                // Save the workbook to a file
                string outputPath = "ZOrderDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Output any unexpected errors
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
