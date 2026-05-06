using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace ImageConversionProgressDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string workbookPath = "input.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(workbookPath);

            // Configure image rendering options (PNG format, one page per sheet)
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                OnePagePerSheet = true
            };

            // Create a renderer for the workbook (lifecycle rule: create)
            WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);

            // Total number of pages that will be rendered
            int totalPages = renderer.PageCount;

            // Collection to hold progress updates (page index -> percentage)
            Dictionary<int, double> progressUpdates = new Dictionary<int, double>();

            // Render each page to an image file and track progress
            for (int pageIndex = 0; pageIndex < totalPages; pageIndex++)
            {
                // Define output image file name
                string imageFile = $"output_page_{pageIndex}.png";

                // Render the current page to the image file (lifecycle rule: save)
                renderer.ToImage(pageIndex, imageFile);

                // Calculate progress percentage
                double progress = ((pageIndex + 1) * 100.0) / totalPages;
                progressUpdates[pageIndex] = progress;

                // Output progress to console (programmatic status update)
                Console.WriteLine($"Rendered page {pageIndex + 1}/{totalPages} to '{imageFile}'. Progress: {progress:0.00}%");
            }

            // Example of retrieving progress programmatically after conversion
            Console.WriteLine("\nConversion progress summary:");
            foreach (var entry in progressUpdates)
            {
                Console.WriteLine($"Page {entry.Key}: {entry.Value:0.00}% completed");
            }

            // Dispose workbook resources
            workbook.Dispose();
        }
    }
}