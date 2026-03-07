using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsConversionProgressDemo
{
    // Implements the page saving callback to track conversion progress.
    public class CustomPageSavingCallback : IPageSavingCallback
    {
        // Called before a page is saved.
        public void PageStartSaving(PageStartSavingArgs args)
        {
            Console.WriteLine($"Starting conversion of page {args.PageIndex + 1} of {args.PageCount}");
        }

        // Called after a page has been saved.
        public void PageEndSaving(PageEndSavingArgs args)
        {
            Console.WriteLine($"Finished conversion of page {args.PageIndex + 1}");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the source XLSX workbook.
            string sourcePath = "input.xlsx";

            // Path for the resulting multi‑page TIFF image.
            string outputPath = "output.tiff";

            // Load the workbook (create/load rule).
            Workbook workbook = new Workbook(sourcePath);

            // Configure image rendering options for TIFF output.
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,
                TiffCompression = TiffCompression.CompressionLZW,
                // Generate multiple pages if the worksheet spans more than one page.
                OnePagePerSheet = false
            };

            // Attach the custom callback to monitor progress.
            options.PageSavingCallback = new CustomPageSavingCallback();

            // Create a renderer for the whole workbook (rendering rule).
            WorkbookRender renderer = new WorkbookRender(workbook, options);

            // Render the workbook to a multi‑page TIFF file (save rule).
            renderer.ToImage(outputPath);

            Console.WriteLine($"Conversion completed. TIFF saved to: {outputPath}");
        }
    }
}