using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeSvgDemo
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                RenderWorksheet();
                Console.WriteLine("Rendering completed.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        private static void RenderWorksheet()
        {
            try
            {
                // Create workbook and fill with sample data
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Item");
                worksheet.Cells["B1"].PutValue("Quantity");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["B3"].PutValue(15);

                // ---------- SVG rendering ----------
                var svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true
                };
                var svgRenderer = new SheetRender(worksheet, svgOptions);
                string svgPath = Path.Combine(Path.GetTempPath(), "worksheet.svg");
                Directory.CreateDirectory(Path.GetDirectoryName(svgPath)!);
                svgRenderer.ToImage(0, svgPath); // page index 0, save to file

                // ---------- PNG rendering ----------
                var pngOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png
                };
                var pngRenderer = new SheetRender(worksheet, pngOptions);
                string pngPath = Path.Combine(Path.GetTempPath(), "worksheet.png");
                pngRenderer.ToImage(0, pngPath); // page index 0, save to file

                Console.WriteLine($"SVG saved to: {svgPath}");
                Console.WriteLine($"PNG saved to: {pngPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to render worksheet: {ex.Message}");
                throw;
            }
        }
    }
}