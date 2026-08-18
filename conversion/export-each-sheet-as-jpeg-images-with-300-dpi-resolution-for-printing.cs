// Title: Export Excel worksheets to 300 DPI JPEG images with Aspose.Cells for .NET
// Description: The sample loads a workbook, creates an output folder, and loops through each worksheet. It configures ImageOrPrintOptions for JPEG format, sets both horizontal and vertical resolution to 300 DPI, forces one page per sheet, renders the sheet with SheetRender, and saves the result as <SheetName>_300dpi.jpg.
// Keywords: Aspose.Cells | C# | .NET | export worksheet to JPEG | 300 DPI | SheetRender | ImageOrPrintOptions | one page per sheet | batch conversion | print‑ready image
// Common Searches: Aspose.Cells export worksheet to JPEG 300 DPI | C# convert Excel sheets to high resolution JPEG | How to render Excel as 300 DPI images using Aspose.Cells | Batch export Excel worksheets to JPEG files .NET | Set DPI when saving Excel as JPEG with Aspose
// Developer Intent: Generate a separate 300 DPI JPEG file for every worksheet in an Excel workbook.
// Use Cases: Produce print‑ready images of each sheet for catalogs or documentation. | Create high‑resolution thumbnails for a web gallery directly from Excel data. | Automate archival‑quality batch conversion of workbook sheets to JPEG.
// AI Prompts: Write C# code that uses Aspose.Cells to export all worksheets of a workbook to 300 DPI JPEG images, one page per sheet. | Show how to modify the example to export only selected worksheets at 600 DPI. | Add robust error handling for missing input files and invalid output directories in the export script.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// The sample loads a workbook, creates an output folder, and loops through each worksheet. It configures ImageOrPrintOptions for JPEG format, sets both horizontal and vertical resolution to 300 DPI, forces one page per sheet, renders the sheet with SheetRender, and saves the result as <SheetName>_300dpi.jpg.
class ExportSheetsToJpeg
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Ensure output directory exists
        string outputDir = "ExportedSheets";
        Directory.CreateDirectory(outputDir);

        // Process each worksheet
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];

            // Set image options: JPEG format, 300 DPI, one page per sheet
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            options.ImageType = ImageType.Jpeg;
            options.HorizontalResolution = 300;
            options.VerticalResolution = 300;
            options.OnePagePerSheet = true;

            // Render the sheet to an image
            SheetRender renderer = new SheetRender(sheet, options);
            string outputPath = Path.Combine(outputDir, $"{sheet.Name}_300dpi.jpg");
            renderer.ToImage(0, outputPath);
        }

        Console.WriteLine("All worksheets have been exported as 300 DPI JPEG images.");
    }
}
