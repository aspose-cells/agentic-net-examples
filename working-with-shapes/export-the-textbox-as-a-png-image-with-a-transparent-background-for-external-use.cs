// Title: Export a TextBox Shape to a Transparent PNG with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a TextBox shape, sets its text, configures ImageOrPrintOptions for PNG with the Transparent flag, and uses the shape's ToImage method to save the textbox as a PNG file that keeps the background transparent.
// Keywords: Aspose.Cells | C# | export textbox PNG | transparent background | shape to image | ImageOrPrintOptions | PNG transparency | ToImage method
// Common Searches: Aspose.Cells export textbox as PNG | transparent PNG from Excel shape C# | how to save Excel textbox with no background using Aspose.Cells | C# Aspose.Cells shape transparent image | export Excel shape to PNG with transparency
// Developer Intent: Generate a PNG file of a worksheet TextBox shape that retains transparency using Aspose.Cells in C#.
// Use Cases: Create UI icons or overlay graphics from spreadsheet textboxes. | Produce watermarks or annotation images for web pages without a background layer. | Export diagram elements for documentation or presentations while keeping the canvas transparent. | Batch‑convert multiple worksheet shapes into transparent PNG assets for design systems.
// AI Prompts: Write C# code that exports all TextBox shapes in a workbook to separate transparent PNG files using Aspose.Cells. | Explain how to adjust DPI and image dimensions in ImageOrPrintOptions while preserving transparency for a TextBox export. | Show how to replace a specific background color with transparency when exporting a shape to PNG with Aspose.Cells. | Provide a script that names each exported PNG file based on the shape's name or index.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a workbook, adds a TextBox shape, sets its text, configures ImageOrPrintOptions for PNG with the Transparent flag, and uses the shape's ToImage method to save the textbox as a PNG file that keeps the background transparent.
class ExportTextboxTransparentPng
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, bottom offset, right offset
        Shape textbox = sheet.Shapes.AddTextBox(2, 1, 0, 0, 100, 200);
        textbox.Text = "Transparent TextBox";

        // Configure image options for PNG with transparent background
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,   // Output format
            Transparent = true           // Enable transparent background
        };

        // Export the textbox shape to a PNG file with transparency
        string outputPath = "textbox.png";
        textbox.ToImage(outputPath, imgOptions);

        Console.WriteLine($"Textbox exported to '{outputPath}' with transparent background.");
    }
}
