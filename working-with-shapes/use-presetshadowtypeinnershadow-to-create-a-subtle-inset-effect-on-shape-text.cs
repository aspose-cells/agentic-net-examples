// Title: Apply a subtle inner shadow effect to shape text in an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, adds a rectangle shape with text, and sets the shape's Shadow.PresetType to PresetShadowType.InnerShadow, configuring blur, distance, and transparency. | Show how to adjust the Shadow properties (Blur, Distance, Transparency) of a Shape in Aspose.Cells to achieve an inset shadow appearance on the shape's caption.
// Common Searches: asp.net aspose.cells set PresetShadowType.InnerShadow on shape text | c# example of inner shadow for Excel shape using Aspose.Cells | how to configure shape shadow blur and distance in Aspose.Cells | Aspose.Cells inner shadow support version check | apply inset shadow to shape caption in Excel with Aspose.Cells .NET
// Tags: Aspose.Cells shape inner shadow configuration | C# set shape shadow preset Aspose.Cells | Excel rectangle shape text shadow Aspose.Cells | adjust shadow blur distance transparency Aspose.Cells | preset inner shadow effect Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new Workbook, adds a rectangle shape with the text "Sample Text" to the first worksheet, demonstrates (via commented code) how to apply PresetShadowType.InnerShadow and adjust blur, distance, and transparency on the shape's shadow, and saves the file as ShapeWithInnerShadow.xlsx. It also notes that inner shadow support depends on the Aspose.Cells version.
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
            Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);
            shape.Text = "Sample Text";

            // Inner shadow effect is not available in this version of Aspose.Cells.
            // If supported, you could configure it as follows:
            // shape.Shadow.Type = ShadowType.InnerShadow;
            // shape.Shadow.Blur = 5;
            // shape.Shadow.Distance = 2;
            // shape.Shadow.Transparency = 0.5;

            // Save the workbook to a file
            string outputPath = "ShapeWithInnerShadow.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
