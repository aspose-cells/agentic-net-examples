// Title: Export a TextBox Shape to a Transparent PNG with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a TextBox shape, configure ImageOrPrintOptions for PNG with transparency, and use Shape.ToImage to save the TextBox as a transparent PNG file. The workbook can also be saved for reference.
// Keywords: Aspose.Cells | .NET | C# | export textbox shape | transparent PNG | Shape.ToImage | ImageOrPrintOptions | Excel shape image | transparent background rendering | save shape as image
// Common Searches: Aspose.Cells export textbox to PNG with transparent background | Shape.ToImage transparent PNG example C# | How to render an Excel shape as a transparent image using Aspose.Cells | Export only a specific shape from a workbook to PNG .NET | Set transparent background when converting Excel shapes to images
// Developer Intent: Generate a PNG file of a worksheet TextBox shape with a transparent background using Aspose.Cells for .NET.
// Use Cases: Create overlay graphics for web pages without background artifacts. | Produce reusable icons of annotated text boxes for documentation or presentations. | Extract individual shape images from Excel workbooks for use in other applications while preserving transparency.
// AI Prompts: Show how to batch‑export all TextBox shapes in a workbook to separate transparent PNG files with Aspose.Cells. | Provide code to export a TextBox shape with a custom DPI and size while keeping the PNG background transparent. | Explain how to export multiple shapes of different types (TextBox, Chart, Picture) to transparent PNG images in one pass.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook, add a TextBox shape, configure ImageOrPrintOptions for PNG with transparency, and use Shape.ToImage to save the TextBox as a transparent PNG file. The workbook can also be saved for reference.
class ExportTextboxAsTransparentPng
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: upper left row, upper left column, row offset, column offset, height, width
        Shape textbox = worksheet.Shapes.AddTextBox(2, 1, 0, 0, 100, 200);
        textbox.Text = "Transparent TextBox";

        // Configure image options for transparent PNG output
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,   // PNG supports transparency
            Transparent = true           // Enable transparent background
        };

        // Export the textbox shape to a PNG file with transparent background
        string outputPath = "textbox_transparent.png";
        textbox.ToImage(outputPath, imgOptions);

        // (Optional) Save the workbook for reference
        workbook.Save("WorkbookWithTextbox.xlsx");

        Console.WriteLine($"Textbox exported to {outputPath} with transparent background.");
    }
}
