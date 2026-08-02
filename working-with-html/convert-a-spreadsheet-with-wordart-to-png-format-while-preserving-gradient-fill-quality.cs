// Title: Export Excel WordArt with Gradient Fill to High‑Resolution PNG using Aspose.Cells for .NET
// Description: Shows how to load an .xlsx workbook that contains WordArt with gradient fill, configure 300 DPI PNG options, and render the first worksheet to a PNG image while preserving gradient quality using Aspose.Cells WorkbookRender in C#.
// Keywords: Aspose.Cells | C# | .NET | Excel to PNG | WordArt export | gradient fill | high DPI image | WorkbookRender | image rendering options | preserve gradient quality
// Common Searches: Aspose.Cells export WordArt to PNG | How to keep gradient colors when converting Excel to PNG | C# render Excel worksheet as high resolution PNG | Save Excel WordArt as image with Aspose.Cells | Convert Excel file with gradient WordArt to PNG .NET
// Developer Intent: Render an Excel worksheet that contains WordArt with gradient fill to a high‑resolution PNG image while maintaining the original gradient appearance.
// Use Cases: Create product catalog thumbnails from Excel templates that include stylized WordArt, ensuring the gradient look is retained. | Generate printable marketing assets by converting Excel reports with decorative WordArt into PNG files for web publishing. | Automate batch conversion of multiple .xlsx files with WordArt to high‑DPI PNG images for archival or distribution.
// AI Prompts: Write C# code using Aspose.Cells to convert an Excel sheet containing WordArt with gradient fill to a 300 DPI PNG while preserving gradient colors. | Explain how ImageOrPrintOptions settings influence the quality of WordArt gradients when rendering to PNG with Aspose.Cells. | Provide a script that processes all .xlsx files in a directory, rendering each first worksheet with WordArt to high‑resolution PNG images.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Shows how to load an .xlsx workbook that contains WordArt with gradient fill, configure 300 DPI PNG options, and render the first worksheet to a PNG image while preserving gradient quality using Aspose.Cells WorkbookRender in C#.
class WordArtToPng
{
    static void Main()
    {
        // Path to the Excel file that contains WordArt with gradient fill
        string sourcePath = "WordArtSample.xlsx";

        // Path where the PNG image will be saved
        string outputPath = "WordArtSample.png";

        // Load the workbook
        Workbook workbook = new Workbook(sourcePath);

        // Set up image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,          // Output format
            HorizontalResolution = 300,         // High DPI to keep gradient quality
            VerticalResolution = 300,
            OnePagePerSheet = true              // Render each sheet as a single page
        };

        // Use WorkbookRender (preserves all drawing objects, including WordArt gradients)
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Render the first sheet (page index 0) to a PNG file
        renderer.ToImage(0, outputPath);

        Console.WriteLine($"WordArt successfully rendered to PNG: {outputPath}");
    }
}
