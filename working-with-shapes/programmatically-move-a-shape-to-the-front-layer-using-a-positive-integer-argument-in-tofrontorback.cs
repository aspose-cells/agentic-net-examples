// Title: How to move a rectangle shape to the front layer in an Excel worksheet using Aspose.Cells for .NET (ToFrontOrBack method)
// Common Searches: Aspose.Cells C# move shape to front layer using ToFrontOrBack | C# example of bringing a rectangle shape to the front in an Excel file with Aspose.Cells | How to set shape order to front in Aspose.Cells workbook programmatically | Using integer argument with ToFrontOrBack to reorder shapes in Aspose.Cells .NET | Aspose.Cells shape layering front back example code
// Tags: Aspose.Cells shape front layer | C# ToFrontOrBack integer parameter | Excel shape ordering Aspose.Cells | add rectangle shape to worksheet Aspose.Cells | programmatic shape layering .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a rectangle shape to the first worksheet, moves that shape to the front layer by calling shape.ToFrontOrBack(1), and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,    // upper left row
                0,    // upper left column
                0,    // top offset (in points)
                100,  // left offset (in points)
                50,   // height (in points)
                150   // width (in points)
            );

            // Move the shape to the front layer.
            // In some Aspose.Cells versions the method expects an int (1 = front, 0 = back)
            shape.ToFrontOrBack(1);

            // Save the workbook
            string outputPath = "Result.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
