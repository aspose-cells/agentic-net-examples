using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgGzipDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and populate it with sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(150);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(250);

                // 2. Set up SVG rendering options (ImageFormat is implicit for SVG)
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true // make SVG fit the viewport
                };

                // 3. Render the first worksheet to an SVG file
                SheetRender renderer = new SheetRender(sheet, svgOptions);
                string svgPath = "worksheet.svg";
                renderer.ToImage(0, svgPath);

                // 4. Read the generated SVG file into a byte array (ensure it exists)
                if (!File.Exists(svgPath))
                    throw new FileNotFoundException($"SVG file not found: {svgPath}");

                byte[] svgBytes = File.ReadAllBytes(svgPath);

                // 5. Compress the SVG bytes using GZip
                string gzPath = "worksheet.svg.gz";
                using (FileStream gzFile = new FileStream(gzPath, FileMode.Create, FileAccess.Write))
                using (GZipStream gzip = new GZipStream(gzFile, CompressionMode.Compress))
                {
                    gzip.Write(svgBytes, 0, svgBytes.Length);
                }

                // 6. Output information about the compression result
                Console.WriteLine($"Original SVG size: {svgBytes.Length} bytes");
                Console.WriteLine($"Compressed GZip size: {new FileInfo(gzPath).Length} bytes");
                Console.WriteLine("SVG file has been compressed and saved as " + gzPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}