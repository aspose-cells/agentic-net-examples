// Title: Insert a rectangle shape with a horizontal red‑to‑blue two‑color gradient fill into an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a rectangle shape to the first worksheet and applies a horizontal red‑to‑blue two‑color gradient using Aspose.Cells. | Create an Excel file where a shape's Fill object is set to Gradient and configured with SetTwoColorGradient for a red‑to‑blue transition in Aspose.Cells .NET. | Save a workbook after configuring a shape's FillType to Gradient and defining a horizontal gradient between Color.Red and Color.Blue.
// Common Searches: Aspose.Cells C# add rectangle shape with horizontal gradient fill | How to set two‑color gradient on a shape in Aspose.Cells .NET | C# example for gradient fill on Excel shape using Aspose.Cells | Create Excel shape with red to blue gradient using Aspose.Cells API
// Tags: Aspose.Cells add rectangle shape | Aspose.Cells gradient fill shape | SetTwoColorGradient Aspose.Cells | horizontal gradient fill C# | Excel shape fill type gradient Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Creates a new workbook, inserts a rectangle shape at row 2 column 2, applies a horizontal red‑to‑blue two‑color gradient fill, and saves the file as ShapeWithGradient.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Add a rectangle shape at row 2, column 2 with width 200 and height 100
            var shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,    // upper left row
                2,    // upper left column
                0,    // top offset (pixels)
                0,    // left offset (pixels)
                100,  // height (pixels)
                200   // width (pixels)
            );

            // Enable gradient fill and set a two‑color gradient (red to blue, horizontal)
            shape.Fill.FillType = FillType.Gradient;
            shape.Fill.SetTwoColorGradient(Color.Red, Color.Blue, GradientStyleType.Horizontal, 0);

            // Determine output path and ensure the directory exists
            string outputPath = "ShapeWithGradient.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
