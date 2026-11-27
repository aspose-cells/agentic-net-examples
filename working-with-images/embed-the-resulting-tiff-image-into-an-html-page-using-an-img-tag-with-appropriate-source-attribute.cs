// Title: Create a TIFF image from an Excel worksheet and embed it in an HTML page with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# render first worksheet page to TIFF and display in a web page | how to generate a 300 dpi LZW‑compressed TIFF from Excel using Aspose.Cells and embed it with an img tag | C# code to convert an Excel sheet to a TIFF file and write an HTML file that references the image | SheetRender ToTiff example with HTML output in Aspose.Cells for .NET | save Excel worksheet as TIFF and create a simple HTML viewer in C#
// Tags: Aspose.Cells SheetRender TIFF export | C# Excel worksheet to TIFF conversion | ImageOrPrintOptions LZW compression settings | HTML img tag embedding generated TIFF | 300 DPI TIFF generation from Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTiffToHtml
{
    // The sample program builds a workbook, fills it with data and formulas, calculates the values, configures ImageOrPrintOptions for a 300 dpi 24‑bit LZW‑compressed TIFF, uses SheetRender to export the first worksheet page to a TIFF file, and then creates an HTML file that embeds the TIFF via an <img> tag.
    class Program
    {
        static void Main(string[] args)
        {
            // Define output file names
            const string tiffFileName = "workbook_page.tiff";
            const string htmlFileName = "workbook_page.html";

            try
            {
                // -------------------------------------------------
                // 1. Create a sample workbook with some data
                // -------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate a few cells with sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["C1"].PutValue("Price");

                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["C2"].PutValue(0.5);

                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(7);
                sheet.Cells["C3"].PutValue(0.3);

                // Add a simple total formula
                sheet.Cells["A4"].PutValue("Total");
                sheet.Cells["B4"].Formula = "=SUM(B2:B3)";
                sheet.Cells["C4"].Formula = "=SUMPRODUCT(B2:B3,C2:C3)";

                // Calculate formulas so that values are available for rendering
                workbook.CalculateFormula();

                // -------------------------------------------------
                // 2. Configure image rendering options for TIFF
                // -------------------------------------------------
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Tiff,                     // Output format
                    TiffCompression = TiffCompression.CompressionLZW, // Reasonable compression
                    HorizontalResolution = 300,                     // DPI
                    VerticalResolution = 300,
                    // Optional: set color depth to 24‑bit for richer colors
                    TiffColorDepth = ColorDepth.Format24bpp
                };

                // -------------------------------------------------
                // 3. Render the first worksheet page to a TIFF file
                // -------------------------------------------------
                SheetRender sheetRender = new SheetRender(sheet, renderOptions);

                // Ensure the worksheet has at least one page
                if (sheetRender.PageCount == 0)
                {
                    Console.WriteLine("No pages were generated for the worksheet. Exiting.");
                    return;
                }

                // Render the first (and only) page to the TIFF file
                sheetRender.ToTiff(tiffFileName);
                sheetRender.Dispose();

                Console.WriteLine($"TIFF image successfully created: {Path.GetFullPath(tiffFileName)}");

                // -------------------------------------------------
                // 4. Generate an HTML page that embeds the TIFF image
                // -------------------------------------------------
                string htmlContent = $@"<!DOCTYPE html>
<html lang=""en"">
<head>
    <meta charset=""UTF-8"">
    <title>Workbook Page as TIFF</title>
    <style>
        body {{ font-family: Arial, Helvetica, sans-serif; margin: 20px; }}
        img {{ border: 1px solid #ccc; max-width: 100%; height: auto; }}
    </style>
</head>
<body>
    <h1>Workbook Page Rendered as TIFF</h1>
    <p>The following image is the rendered first page of the workbook:</p>
    <img src=""{tiffFileName}"" alt=""Workbook page rendered as TIFF"" />
</body>
</html>";

                File.WriteAllText(htmlFileName, htmlContent);
                Console.WriteLine($"HTML file successfully created: {Path.GetFullPath(htmlFileName)}");
                Console.WriteLine("Open the HTML file in a browser to view the embedded TIFF image.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred during processing:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
