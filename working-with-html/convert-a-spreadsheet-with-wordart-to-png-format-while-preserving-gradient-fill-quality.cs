// Title: Export Excel WordArt with Gradient Fill to PNG using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, adds a WordArt shape with a two‑color horizontal gradient, sets 300 DPI PNG rendering options, and saves the worksheet as a single high‑quality PNG while keeping the gradient intact.
// Keywords: Aspose.Cells | C# | Excel WordArt export | gradient fill PNG | high DPI image rendering | WorkbookRender | ImageOrPrintOptions | preserve WordArt colors | convert Excel to PNG | Office automation
// Common Searches: Aspose.Cells export WordArt to PNG | C# render Excel gradient WordArt as image | high resolution PNG from Excel workbook | preserve WordArt gradient when converting to PNG | how to save Excel sheet with WordArt as PNG
// Developer Intent: Generate a PNG image from an Excel sheet that contains gradient‑filled WordArt, ensuring the gradient is rendered accurately.
// Use Cases: Create web‑ready graphics from Excel‑based marketing designs without losing gradient effects. | Produce printable, high‑DPI PNG assets from reports that include WordArt for inclusion in PDFs or slide decks. | Automate batch conversion of multiple workbooks containing WordArt into PNG files for archival or distribution.
// AI Prompts: Show how to export each worksheet to its own PNG file while preserving WordArt gradients. | Demonstrate applying a multi‑stop gradient to WordArt before rendering to PNG with Aspose.Cells. | Explain how to change DPI and image size for optimal PNG output of gradient WordArt using ImageOrPrintOptions.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, adds a WordArt shape with a two‑color horizontal gradient, sets 300 DPI PNG rendering options, and saves the worksheet as a single high‑quality PNG while keeping the gradient intact.
class ConvertWordArtToPng
{
    static void Main()
    {
        // Path to the source Excel file that contains WordArt (or create a new one)
        string sourcePath = "WordArtWorkbook.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(sourcePath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the shape collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add a WordArt shape with a preset gradient style (WordArtStyle7)
        // Parameters: style, text, topRow, top (pixels), leftColumn, left (pixels), height (pixels), width (pixels)
        Shape wordArt = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,
            "Gradient WordArt",
            2,          // topRow
            0,          // top offset in pixels
            2,          // leftColumn
            0,          // left offset in pixels
            100,        // height in pixels
            400);       // width in pixels

        // Ensure the fill type is set to gradient and define a custom two‑color gradient
        wordArt.Fill.FillType = FillType.Gradient;
        GradientFill gradientFill = wordArt.Fill.GradientFill;
        if (gradientFill != null)
        {
            // Create a horizontal gradient from blue to light blue
            gradientFill.SetTwoColorGradient(
                Color.Blue,          // first color
                Color.LightBlue,     // second color
                GradientStyleType.Horizontal,
                1);                  // variant
        }

        // Configure image rendering options for high‑quality PNG output
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,          // PNG format
            HorizontalResolution = 300,         // 300 DPI horizontal
            VerticalResolution = 300,           // 300 DPI vertical
            OnePagePerSheet = true              // Render each sheet as a single page
        };

        // Render the entire workbook to a PNG file
        string outputPath = "WordArtRendered.png";
        WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);
        renderer.ToImage(outputPath); // Uses WorkbookRender.ToImage(string) overload

        Console.WriteLine($"Workbook with WordArt successfully rendered to PNG: {outputPath}");
    }
}
