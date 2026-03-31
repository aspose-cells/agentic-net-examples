using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class ProgressImageConversion
{
    static void Main()
    {
        // Load the source XLSX workbook
        string sourcePath = "input.xlsx";
        Workbook workbook = new Workbook(sourcePath); // uses Workbook(string) constructor

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,               // output image format
            PageSavingCallback = new ConsolePageSavingCallback() // attach progress callback
        };

        // Create a renderer for the entire workbook
        WorkbookRender renderer = new WorkbookRender(workbook, options); // uses WorkbookRender(Workbook, ImageOrPrintOptions) constructor

        // Render each page of the workbook to a separate image file while monitoring progress
        for (int i = 0; i < renderer.PageCount; i++)
        {
            string outputFile = $"output_page_{i}.png";
            renderer.ToImage(i, outputFile); // renders page i to the specified file
        }

        // Release resources
        renderer.Dispose();
        workbook.Dispose();
    }

    // Custom callback implementation to monitor page saving progress
    private class ConsolePageSavingCallback : IPageSavingCallback
    {
        public void PageStartSaving(PageStartSavingArgs args)
        {
            Console.WriteLine($"Starting to save page {args.PageIndex}");
        }

        public void PageEndSaving(PageEndSavingArgs args)
        {
            Console.WriteLine($"Finished saving page {args.PageIndex}");
        }
    }
}