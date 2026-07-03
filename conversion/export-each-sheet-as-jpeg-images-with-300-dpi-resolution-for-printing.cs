using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class ExportSheetsToJpeg
{
    static void Main()
    {
        // Path to the source Excel file
        string inputPath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(inputPath);

        // Configure image options: JPEG format with 300 DPI resolution
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Jpeg;          // Set output format to JPEG
        options.HorizontalResolution = 300;          // 300 DPI horizontally
        options.VerticalResolution = 300;            // 300 DPI vertically
        options.OnePagePerSheet = true;              // Render each sheet as a single page

        // Create output directory for the images
        string outputDir = "SheetImages";
        Directory.CreateDirectory(outputDir);

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Initialize SheetRender with the current worksheet and image options
            SheetRender renderer = new SheetRender(sheet, options);

            // Render each page of the sheet (usually one page because of OnePagePerSheet)
            for (int pageIndex = 0; pageIndex < renderer.PageCount; pageIndex++)
            {
                // Build a filename that includes the sheet name and page number
                string fileName = Path.Combine(
                    outputDir,
                    $"{sheet.Name}_Page{pageIndex + 1}.jpg");

                // Export the page to a JPEG image file
                renderer.ToImage(pageIndex, fileName);
            }

            // Release resources used by the renderer
            renderer.Dispose();
        }

        Console.WriteLine("All sheets have been exported as 300 DPI JPEG images.");
    }
}