using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,          // Output format
            OnePagePerSheet = true,             // Render each page separately
            PageSavingCallback = new ConsolePageSavingCallback() // Progress callback
        };

        // Create a renderer for the workbook
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Prepare output directory
        string outputDir = "output_images";
        Directory.CreateDirectory(outputDir);

        // Render each page to an image file while displaying progress
        for (int i = 0; i < renderer.PageCount; i++)
        {
            string filePath = Path.Combine(outputDir, $"page_{i + 1}.png");
            renderer.ToImage(i, filePath);
        }

        Console.WriteLine("Image conversion completed.");
    }

    // Implementation of IPageSavingCallback to report progress to the console
    private class ConsolePageSavingCallback : IPageSavingCallback
    {
        public void PageStartSaving(PageStartSavingArgs args)
        {
            Console.WriteLine($"Starting to render page {args.PageIndex + 1} of {args.PageCount}");
        }

        public void PageEndSaving(PageEndSavingArgs args)
        {
            Console.WriteLine($"Finished rendering page {args.PageIndex + 1}");
        }
    }
}