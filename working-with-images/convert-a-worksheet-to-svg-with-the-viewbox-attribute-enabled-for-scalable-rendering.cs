using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgExport
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Quantity");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(150);
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["B3"].PutValue(250);

                // Set SVG rendering options (FitToViewPort enables viewBox)
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true
                };

                // Render the worksheet to SVG
                SheetRender renderer = new SheetRender(worksheet, svgOptions);
                string outputPath = "WorksheetOutput.svg";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Export the first page (index 0) as SVG
                renderer.ToImage(0, outputPath);

                Console.WriteLine("Worksheet has been exported to SVG with FitToViewPort enabled.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}