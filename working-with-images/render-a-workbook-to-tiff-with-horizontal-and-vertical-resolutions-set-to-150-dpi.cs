using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTiffRenderDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates an empty workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data to demonstrate rendering
            sheet.Cells["A1"].PutValue("Aspose.Cells TIFF Rendering Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Configure image rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,          // Set output format to TIFF
                HorizontalResolution = 150,          // Set horizontal DPI
                VerticalResolution = 150,            // Set vertical DPI
                OnePagePerSheet = true               // Render the whole sheet on one page
            };

            // Create a workbook renderer with the specified options
            WorkbookRender renderer = new WorkbookRender(workbook, options);

            // Render the entire workbook to a multi‑page TIFF file
            string outputPath = "RenderedWorkbook.tiff";
            renderer.ToImage(outputPath);

            Console.WriteLine($"Workbook successfully rendered to TIFF at: {outputPath}");
        }
    }
}