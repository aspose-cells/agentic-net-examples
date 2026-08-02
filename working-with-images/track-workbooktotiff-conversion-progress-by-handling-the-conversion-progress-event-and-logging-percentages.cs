using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class WorkbookToTiffProgressDemo
{
    static void Main()
    {
        // Create a new workbook and populate it with enough data to span multiple pages
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        for (int i = 0; i < 300; i++)
        {
            sheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Configure image options for TIFF rendering
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,
            TiffCompression = TiffCompression.CompressionLZW,
            OnePagePerSheet = false // generate a multi‑page TIFF
        };

        // Attach the progress callback
        options.PageSavingCallback = new ProgressCallback();

        // Render the worksheet to a multi‑page TIFF file
        SheetRender renderer = new SheetRender(sheet, options);
        renderer.ToTiff("WorkbookToTiffWithProgress.tiff");
    }

    // Callback that logs the conversion progress as a percentage
    private class ProgressCallback : IPageSavingCallback
    {
        public void PageStartSaving(PageStartSavingArgs args)
        {
            int totalPages = args.PageCount;
            int currentPage = args.PageIndex + 1; // convert zero‑based index to 1‑based
            double percent = (double)currentPage / totalPages * 100;
            Console.WriteLine($"Saving page {currentPage}/{totalPages} ({percent:0.##}%)");
        }

        public void PageEndSaving(PageEndSavingArgs args)
        {
            // Optional: indicate when a page has finished rendering
            Console.WriteLine($"Finished page {args.PageIndex + 1}");
        }
    }
}