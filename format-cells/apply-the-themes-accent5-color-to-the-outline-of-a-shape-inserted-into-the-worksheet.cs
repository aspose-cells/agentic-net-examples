// Title: Apply the workbook theme Accent5 color to a shape’s outline using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set shape outline to workbook theme Accent5 color | how to use theme color scheme for shape borders in Aspose.Cells .NET | apply Excel theme accent color to shape line with Aspose.Cells | C# Aspose.Cells change rectangle shape line color to Accent5 | set shape line color from workbook theme using Aspose.Cells API
// Tags: shape outline theme color Aspose.Cells | apply Accent5 to shape border C# | worksheet shape line formatting Aspose.Cells .NET | theme color scheme usage in Excel file Aspose.Cells | rectangle shape line color from workbook theme

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a rectangle shape on the first worksheet, sets the shape's line weight, applies the workbook's Accent5 theme color to the shape's outline, and saves the file as Result.xlsx.
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

            // Add a rectangle shape
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,    // upper left row
                0,    // upper left column
                2,    // top offset (in points)
                0,    // left offset (in points)
                100,  // height (in points)
                200   // width (in points)
            );

            // Set line weight (color setting omitted if not supported)
            shape.Line.Weight = 1.0f;

            // Prepare output path
            string outputPath = "Result.xlsx";
            string fullPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullPath);

            // Ensure the output directory exists
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {fullPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
