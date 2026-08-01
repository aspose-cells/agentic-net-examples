// Title: Apply Outer Bottom Preset Shadow to Shape Text and Export as PNG – Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts a rectangle auto‑shape with the text "Shadow Demo", sets the text's shadow to the OffsetBottom preset, saves the file, and renders the shape to a PNG image using Shape.ToImage with custom image options.
// Keywords: Aspose.Cells C# shape shadow | PresetShadowType.OffsetBottom | shape to PNG | Shape.ToImage example | auto shape rectangle Aspose.Cells
// Common Searches: Aspose.Cells add outer bottom shadow to shape text | C# capture shape as PNG with Aspose.Cells | How to use Shape.ToImage in Aspose.Cells | PresetShadowType OffsetBottom sample code
// Developer Intent: Add an outer‑bottom preset shadow to a shape’s text and generate a PNG screenshot of that shape using Aspose.Cells for .NET.
// Use Cases: Produce diagram assets with shadowed labels for documentation or slide decks. | Create thumbnail previews of styled shapes for a reporting dashboard. | Automate generation of branded graphics where visual effects are applied before exporting as image files.
// AI Prompts: Show how to change the preset shadow to OffsetTop and export the shape as a JPEG with Aspose.Cells. | Provide code that captures screenshots of every shape on a worksheet and merges them into one image. | Explain how to adjust shadow offset, blur radius, and color for shape text in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a workbook, inserts a rectangle auto‑shape with the text "Shadow Demo", sets the text's shadow to the OffsetBottom preset, saves the file, and renders the shape to a PNG image using Shape.ToImage with custom image options.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape with some text
            Shape shape = sheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 2, 2, 2, 2, 200, 100);
            shape.Text = "Shadow Demo";

            // Apply an outer bottom preset shadow to the shape's text
            shape.TextOptions.Shadow.PresetType = PresetShadowType.OffsetBottom;

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("ShadowDemo.xlsx", SaveFormat.Xlsx);

            // Set image options for PNG format
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Png,
                HorizontalResolution = 96,
                VerticalResolution = 96
            };

            // Capture a screenshot of the shape and save it as an image
            shape.ToImage("ShadowDemoShape.png", imgOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
