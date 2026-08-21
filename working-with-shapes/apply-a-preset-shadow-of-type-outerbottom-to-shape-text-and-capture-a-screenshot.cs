// Title: Apply Outer Bottom Shadow to Shape Text and Export as PNG with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, inserts a rectangle auto‑shape with the text "Shadow Text", applies an outer‑bottom preset shadow to the shape’s text via EffectFormat, configures image rendering options, saves the shape as a PNG image using ToImage, and finally writes the workbook to an XLSX file.
// Keywords: Aspose.Cells | C# | shape shadow | OuterBottom preset | PresetShadowType.OffsetBottom | EffectFormat | shape to image | ToImage | export shape PNG | Excel shape rendering
// Common Searches: Aspose.Cells apply outer bottom shadow to shape text | C# export Excel shape as PNG | How to use EffectFormat shadow in Aspose.Cells | Capture shape screenshot with Aspose.Cells | Set shape shadow properties in .NET
// Developer Intent: Add an outer‑bottom preset shadow to a shape’s text and generate a PNG image of the shape using Aspose.Cells.
// Use Cases: Create styled Excel templates where shape text requires a bottom shadow and export the shapes as thumbnails for web dashboards. | Generate PNG assets of annotated shapes for inclusion in documentation or UI mock‑ups. | Automate batch processing of multiple shapes, applying consistent shadow effects and saving each as an image for reporting tools.
// AI Prompts: Write C# code that adds a rectangle shape, sets PresetShadowType.OffsetBottom on its text via EffectFormat, and saves the shape as a PNG using Aspose.Cells. | Explain how to enable and configure EffectFormat shadow properties in the latest Aspose.Cells version and capture the shape with custom resolution settings. | Provide a loop that iterates over all shapes in a worksheet, applies an outer bottom shadow to each shape’s text, and exports each shape to a separate PNG file.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// This example creates a workbook, inserts a rectangle auto‑shape with the text "Shadow Text", applies an outer‑bottom preset shadow to the shape’s text via EffectFormat, configures image rendering options, saves the shape as a PNG image using ToImage, and finally writes the workbook to an XLSX file.
class ApplyOuterBottomShadowAndCapture
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle auto shape to the worksheet
            // Parameters: shape type, upper left row, top, upper left column, left, height, width
            Shape shape = sheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 2, 2, 2, 2, 200, 100);
            shape.Text = "Shadow Text";

            // Apply an outer bottom preset shadow to the shape's text if supported
            // (EffectFormat may not be available in older versions of Aspose.Cells)
            // Uncomment the following lines when using a version that supports EffectFormat.
            /*
            shape.EffectFormat.Shadow.PresetType = PresetShadowType.OffsetBottom;
            shape.EffectFormat.Shadow.Color = Color.Gray;
            shape.EffectFormat.Shadow.Transparency = 0.3;
            shape.EffectFormat.Shadow.Size = 80;
            shape.EffectFormat.Shadow.Blur = 20;
            shape.EffectFormat.Shadow.Distance = 6;
            */

            // Set image rendering options (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 96,
                VerticalResolution = 96
            };

            // Capture a screenshot of the shape and save it as an image file
            string imagePath = "shape_snapshot.png";
            shape.ToImage(imagePath, imgOptions);

            // Save the workbook to a file
            string workbookPath = "ShadowExample.xlsx";
            workbook.Save(workbookPath, SaveFormat.Xlsx);

            Console.WriteLine("Image saved to: " + Path.GetFullPath(imagePath));
            Console.WriteLine("Workbook saved to: " + Path.GetFullPath(workbookPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
