using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class WorksheetToImageDemo
    {
        public static void Run()
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Path for the output image (PNG format)
            string outputImagePath = "worksheet_image.png";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure image rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                // Ensure each worksheet is rendered as a single page
                OnePagePerSheet = true
            };

            // Create a SheetRender instance
            SheetRender sheetRender = new SheetRender(worksheet, options);
            // Render the first (and only) page of the worksheet to an image file
            sheetRender.ToImage(0, outputImagePath);
            Console.WriteLine($"Worksheet successfully rendered to image: {outputImagePath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetToImageDemo.Run();
        }
    }
}