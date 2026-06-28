using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class SplitTiffPages
{
    static void Main()
    {
        // Path to the source Excel workbook that will be rendered as a multi‑page TIFF
        string workbookPath = "input.xlsx";

        // Directory where individual single‑page TIFF files will be saved
        string outputFolder = "output_pages";
        Directory.CreateDirectory(outputFolder);

        // Load the workbook
        Workbook workbook = new Workbook(workbookPath);

        // Configure rendering options for TIFF output
        ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,               // Set output format to TIFF
            TiffCompression = TiffCompression.CompressionLZW,
            OnePagePerSheet = false                    // Allow multiple pages per sheet
        };

        // Create a WorkbookRender instance (uses the provided rule)
        WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);

        // Iterate through each rendered page and save it as a separate TIFF file
        for (int pageIndex = 0; pageIndex < renderer.PageCount; pageIndex++)
        {
            string pageFile = Path.Combine(outputFolder, $"page_{pageIndex + 1}.tiff");
            // Save the specific page to a file (uses ToImage(int, string) rule)
            renderer.ToImage(pageIndex, pageFile);
            Console.WriteLine($"Saved page {pageIndex + 1} to {pageFile}");
        }

        // Release resources
        renderer.Dispose();
    }
}