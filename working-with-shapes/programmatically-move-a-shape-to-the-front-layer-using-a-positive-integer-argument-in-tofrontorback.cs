// Title: How to move a rectangle shape to the front layer in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape order to front in Excel worksheet | move rectangle shape to front layer programmatically with Aspose.Cells | using ToFrontOrBack(1) to bring a shape forward in .NET Excel file | change z-order of shapes in Aspose.Cells workbook C# | bring a shape to the top of other objects in Excel using Aspose.Cells API
// Tags: Aspose.Cells shape z-order manipulation | ToFrontOrBack method C# | add rectangle shape Aspose.Cells | Excel worksheet shape layering .NET | programmatic shape ordering Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads (or creates) an Excel workbook, adds a rectangle shape to the first worksheet, calls shape.ToFrontOrBack(1) to move the shape to the front layer, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Get the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape directly (AddShape returns a Shape object in newer API versions).
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                1,   // upper left column
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                100, // height (pixels)
                50   // width (pixels)
            );

            // Bring the shape to the front layer (1 = front, 0 = back).
            shape.ToFrontOrBack(1);

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
