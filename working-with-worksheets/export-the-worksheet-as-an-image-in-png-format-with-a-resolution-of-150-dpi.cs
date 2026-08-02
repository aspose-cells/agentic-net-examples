using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class ExportWorksheetToPng
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data to the worksheet
        sheet.Cells["A1"].PutValue("Sample");
        sheet.Cells["B1"].PutValue("Data");

        // Set up image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;          // Export as PNG
        options.HorizontalResolution = 150;         // 150 DPI horizontal
        options.VerticalResolution = 150;           // 150 DPI vertical
        options.OnePagePerSheet = true;             // Render the whole sheet on one page

        // Create a SheetRender instance with the worksheet and options
        SheetRender renderer = new SheetRender(sheet, options);

        // Prepare output directory and file path
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "worksheet.png");

        // Render the first (and only) page of the worksheet to a PNG file
        renderer.ToImage(0, outputPath);

        Console.WriteLine($"Worksheet exported to PNG at 150 DPI: {outputPath}");
    }
}