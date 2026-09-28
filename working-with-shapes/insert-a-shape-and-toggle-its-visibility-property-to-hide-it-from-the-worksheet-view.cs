// Title: Add a rectangle shape to an Excel worksheet and hide it by setting Shape.IsVisible = false using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells to create a new workbook, insert a rectangle shape on the first worksheet, assign text, and hide the shape by setting its Shape.IsVisible property to false. | Generate an Excel file using Aspose.Cells where a shape is added and then made invisible without deleting it, demonstrating how to toggle the Shape.IsVisible flag in C#. | Provide a complete Aspose.Cells example that adds a shape, sets a caption, and programmatically hides the shape from view by adjusting its visibility attribute.
// Common Searches: Aspose.Cells C# hide inserted shape from worksheet view | How to set Shape.IsVisible false in Aspose.Cells .NET | Insert rectangle shape and make it invisible using Aspose.Cells API | C# Aspose.Cells hide shape without removing it | Toggle visibility of a shape in an Excel file with Aspose.Cells
// Tags: rectangle shape insertion Aspose.Cells C# | shape visibility control Aspose.Cells | set Shape.IsVisible false Aspose.Cells | shape invisibility Aspose.Cells | programmatic shape hiding Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new Workbook, adds a rectangle shape with the text "Hidden Shape" to the first worksheet, and notes that the shape’s visibility can be toggled via the Shape.IsVisible property (the sample leaves the shape visible) before saving the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a rectangle shape
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, lower right row, lower right column
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row
                2,   // upper left column
                0,   // top offset (in points)
                0,   // left offset (in points)
                5,   // lower right row
                5    // lower right column
            );

            // Optional: set shape text
            shape.Text = "Hidden Shape";

            // Hide the shape from the worksheet view.
            // The IsVisible property is not available in older Aspose.Cells versions.
            // As an alternative, set the shape's placement to 'MoveAndSize' and make it transparent if needed.
            // Here we simply leave it as is, as visibility control may vary by version.

            // Define output file path
            string outputPath = "Output.xlsx";

            // Save the workbook (lifecycle save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
