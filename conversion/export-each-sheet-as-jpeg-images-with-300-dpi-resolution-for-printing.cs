// Title: Export every Excel worksheet to a 300 DPI JPEG image using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, configures ImageOrPrintOptions for JPEG at 300 DPI, and saves each worksheet as an individual .jpg file with Aspose.Cells. | Show how to use SheetRender together with ImageOrPrintOptions to produce high‑resolution JPEG images for all sheets in a workbook.
// Common Searches: how to save each sheet of an Excel workbook as a 300 dpi JPEG using Aspose.Cells C# | Aspose.Cells C# export workbook worksheets to high resolution JPEG images | C# convert Excel worksheets to separate JPEG files with 300 DPI resolution | using ImageOrPrintOptions to set JPEG resolution for Excel sheet rendering in .NET
// Tags: Aspose.Cells export worksheets to JPEG 300dpi | ImageOrPrintOptions JPEG resolution setting | SheetRender save sheet as image .NET | C# high‑resolution Excel to JPEG conversion | batch export Excel sheets as JPEG files

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, configures ImageOrPrintOptions for JPEG output at 300 DPI, iterates through each worksheet, and uses SheetRender to render the first page of every sheet to a separate JPEG file named after the sheet, storing the images in a designated output folder.
class ExportSheetsToJpeg
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Ensure the output directory exists
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Configure image options: JPEG format with 300 DPI resolution
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Jpeg;          // Set output format to JPEG
        options.HorizontalResolution = 300;          // 300 DPI horizontally
        options.VerticalResolution = 300;            // 300 DPI vertically
        options.OnePagePerSheet = true;              // One image per sheet

        // Iterate through all worksheets and export each as a JPEG image
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];

            // Create a renderer for the current worksheet
            SheetRender renderer = new SheetRender(sheet, options);

            // Build a safe file name based on the sheet name
            string safeSheetName = sheet.Name;
            foreach (char c in Path.GetInvalidFileNameChars())
                safeSheetName = safeSheetName.Replace(c, '_');

            string outputPath = Path.Combine(outputDir, $"{safeSheetName}.jpg");

            // Render the first (and only) page of the sheet to the JPEG file
            renderer.ToImage(0, outputPath);
        }

        Console.WriteLine("All worksheets have been exported as 300 DPI JPEG images.");
    }
}
