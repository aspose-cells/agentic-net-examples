// Title: Create a Base64‑encoded PNG data URI from an Aspose.Cells worksheet and embed it in a CSS background using C#
// Common Searches: how to render an Excel worksheet to PNG and get a data URI in C# with Aspose.Cells | C# generate CSS background-image from Aspose.Cells worksheet image | convert Aspose.Cells SheetRender output to base64 string for inline CSS | embed Excel sheet snapshot as data:image/png in stylesheet using Aspose.Cells | Aspose.Cells render worksheet to memory stream and create data URI for web page
// Tags: Aspose.Cells SheetRender PNG data URI | C# convert rendered worksheet to base64 | inline CSS background-image from Excel image | Excel worksheet image to data URI Aspose.Cells | render worksheet to memory stream .NET | base64 PNG for CSS background C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsDataUriDemo
{
    // The example creates a workbook, uses Aspose.Cells SheetRender to render the first worksheet to a PNG image in a memory stream, encodes the PNG bytes to a Base64 data URI, and builds a CSS snippet that embeds the data URI as a background-image.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Create a sample workbook with some data
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate a few cells so the rendered image is not empty
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1.20);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(0.80);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(2.50);

                // Optional: adjust column widths for better appearance
                sheet.AutoFitColumn(0);
                sheet.AutoFitColumn(1);

                // ------------------------------------------------------------
                // 2. Render the worksheet to a PNG image in memory
                // ------------------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,      // PNG for the data URI
                    OnePagePerSheet = true,         // Render the whole sheet as a single page
                    Transparent = true              // Keep background transparent (optional)
                };

                SheetRender renderer = new SheetRender(sheet, imgOptions);

                // ------------------------------------------------------------
                // 3. Convert the rendered image to a Base64‑encoded PNG data URI
                // ------------------------------------------------------------
                string dataUri;
                using (MemoryStream ms = new MemoryStream())
                {
                    // Render the first page directly into the memory stream
                    renderer.ToImage(0, ms);
                    byte[] pngBytes = ms.ToArray();

                    // Encode the byte array to Base64 and build the data URI
                    string base64 = Convert.ToBase64String(pngBytes);
                    dataUri = $"data:image/png;base64,{base64}";
                }

                // ------------------------------------------------------------
                // 4. Produce a CSS snippet that uses the data URI as a background
                // ------------------------------------------------------------
                string css = $@".worksheet-background {{
    width: 100%;
    height: auto;
    background-image: url(""{dataUri}"");
    background-repeat: no-repeat;
    background-size: contain;
}}";

                // Output the CSS to the console
                Console.WriteLine("=== CSS with embedded worksheet image ===");
                Console.WriteLine(css);
                Console.WriteLine("=== End of CSS ===");

                // (Optional) Save the CSS to a file for verification
                try
                {
                    File.WriteAllText("worksheet-background.css", css);
                    Console.WriteLine("CSS file saved as 'worksheet-background.css'.");
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Failed to write CSS file: {ex.Message}");
                }

                // (Optional) Save the PNG image to disk for visual comparison
                try
                {
                    renderer.ToImage(0, "worksheet.png");
                    Console.WriteLine("Rendered PNG saved as 'worksheet.png'.");
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Failed to save PNG image: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
