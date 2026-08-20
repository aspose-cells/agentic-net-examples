// Title: Export a Worksheet to PNG at 150 DPI with Aspose.Cells (C#)
// Description: Demonstrates how to create or load a workbook, set ImageOrPrintOptions for PNG format with 150 DPI horizontal and vertical resolution, enable OnePagePerSheet, and use SheetRender to save the first page as a high‑resolution PNG file.
// Keywords: Aspose.Cells | C# | Export worksheet to PNG | 150 DPI | ImageOrPrintOptions | SheetRender | OnePagePerSheet | Excel to image | high resolution PNG | render Excel sheet
// Common Searches: Aspose.Cells export worksheet PNG 150 DPI C# | set image resolution Aspose.Cells C# | render entire Excel sheet as PNG Aspose.Cells | ImageOrPrintOptions DPI setting example | SheetRender save worksheet as PNG
// Developer Intent: Generate a PNG image of a worksheet at 150 DPI using Aspose.Cells in C#.
// Use Cases: Create high‑resolution PNG snapshots of Excel reports for web dashboards. | Produce printable PNG assets for documentation, PDFs, or slide decks. | Automate batch conversion of multiple worksheets to 150 DPI PNG for archival or distribution.
// AI Prompts: Write C# code that uses Aspose.Cells to export a specific worksheet to a PNG image with 150 DPI, including all required option settings. | Explain how to configure ImageOrPrintOptions for DPI, image format, and page layout when rendering an Excel sheet with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Demonstrates how to create or load a workbook, set ImageOrPrintOptions for PNG format with 150 DPI horizontal and vertical resolution, enable OnePagePerSheet, and use SheetRender to save the first page as a high‑resolution PNG file.
class ExportWorksheetToPng
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // create
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Sample Text");
        sheet.Cells["B2"].PutValue(12345);

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;          // PNG format
        options.HorizontalResolution = 150;        // 150 DPI horizontal
        options.VerticalResolution = 150;          // 150 DPI vertical
        options.OnePagePerSheet = true;            // render whole sheet on one page

        // Create a SheetRender instance with the worksheet and options
        SheetRender renderer = new SheetRender(sheet, options);

        // Define output file path
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Worksheet.png");

        // Render the first (and only) page to the PNG file
        renderer.ToImage(0, outputPath); // export

        Console.WriteLine($"Worksheet exported to PNG at 150 DPI: {outputPath}");
    }
}
