using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class WorksheetPreviewRenderer
    {
        public static void Run()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure image rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                OnePagePerSheet = true,
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Create a SheetRender instance
            SheetRender sheetRender = new SheetRender(worksheet, options);

            // Ensure output directory exists
            string outputDir = "output";
            Directory.CreateDirectory(outputDir);

            // Render the first page of the worksheet to an image file
            string imagePath = Path.Combine(outputDir, "worksheet_preview.png");
            sheetRender.ToImage(0, imagePath);

            Console.WriteLine($"Worksheet preview rendered successfully to: {imagePath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetPreviewRenderer.Run();
        }
    }
}