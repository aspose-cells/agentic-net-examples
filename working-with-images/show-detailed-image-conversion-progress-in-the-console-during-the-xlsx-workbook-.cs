using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and populate it with enough rows to generate multiple pages
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        for (int i = 0; i < 200; i++)
        {
            sheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Configure page setup so that rendering will produce several pages
        sheet.PageSetup.FitToPagesWide = 1; // one page wide
        sheet.PageSetup.FitToPagesTall = 0; // unlimited pages tall

        // Set image rendering options and attach a console progress callback
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            PageSavingCallback = new ConsolePageSavingCallback()
        };

        // Create the renderer
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        Console.WriteLine($"Total pages to render: {renderer.PageCount}");

        // Render each page to a separate PNG file while the callback reports progress
        for (int i = 0; i < renderer.PageCount; i++)
        {
            string fileName = $"page_{i + 1}.png";
            renderer.ToImage(i, fileName);
        }

        Console.WriteLine("Image conversion completed.");
    }

    // Implementation of IPageSavingCallback that writes progress to the console
    private class ConsolePageSavingCallback : IPageSavingCallback
    {
        public void PageStartSaving(PageStartSavingArgs args)
        {
            Console.WriteLine($"Start saving page {args.PageIndex + 1} of {args.PageCount}");
        }

        public void PageEndSaving(PageEndSavingArgs args)
        {
            Console.WriteLine($"Finished saving page {args.PageIndex + 1}");
        }
    }
}