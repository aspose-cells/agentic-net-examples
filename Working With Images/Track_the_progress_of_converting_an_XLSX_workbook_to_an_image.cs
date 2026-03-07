using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Load the workbook from file
        Workbook workbook = new Workbook(sourcePath);

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            // Render as multi‑page TIFF
            ImageType = ImageType.Tiff,
            TiffCompression = TiffCompression.CompressionLZW,
            // Allow multiple pages per sheet to demonstrate paging
            OnePagePerSheet = false
        };

        // Attach a custom callback to track page rendering progress
        options.PageSavingCallback = new ImageConversionProgressCallback();

        // Create a renderer for the whole workbook
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Output image file path
        string outputPath = "output.tiff";

        // Render the workbook to the image file; the callback will be invoked for each page
        renderer.ToImage(outputPath);

        // Display final information
        Console.WriteLine($"Rendering completed. Total pages rendered: {renderer.PageCount}");
    }

    // Implementation of IPageSavingCallback to monitor page start and end events
    private class ImageConversionProgressCallback : IPageSavingCallback
    {
        // Called before a page is rendered
        public void PageStartSaving(PageStartSavingArgs args)
        {
            Console.WriteLine($"Starting to render page {args.PageIndex + 1} of {args.PageCount}");
        }

        // Called after a page has been rendered
        public void PageEndSaving(PageEndSavingArgs args)
        {
            Console.WriteLine($"Finished rendering page {args.PageIndex + 1}");
        }
    }
}