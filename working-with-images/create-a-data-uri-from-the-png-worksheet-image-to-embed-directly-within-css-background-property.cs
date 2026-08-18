// Title: C# – Convert Aspose.Cells Worksheet to PNG Data URI for CSS Background
// Description: Render the first worksheet of an Aspose.Cells workbook to a PNG image in memory, encode it as a Base64 data URI, and embed the URI directly in a CSS background‑image rule without creating a separate file.
// Keywords: Aspose.Cells PNG data URI | C# render worksheet to image | base64 image for CSS background | inline image data URI Aspose | .NET Excel to PNG | embed Excel snapshot in CSS | memory stream image conversion
// Common Searches: Aspose.Cells render worksheet to PNG data URI C# | how to embed Excel sheet image in CSS background | convert workbook page to base64 string for inline CSS | C# generate data:image/png from Aspose.Cells | inline PNG image from Excel without saving file
// Developer Intent: Create a PNG data URI from a rendered worksheet and use it as a CSS background image.
// Use Cases: Display an Excel worksheet snapshot on a web page without hosting an external image file. | Include worksheet graphics in email templates using inline CSS to avoid attachment limits. | Generate self‑contained HTML reports where worksheet images are embedded directly in the markup.
// AI Prompts: Write a C# method that takes an Aspose.Cells Workbook and returns a data:image/png;base64 string for the first worksheet. | Show how to render a worksheet to a MemoryStream, convert it to Base64, and build a CSS background rule in C#. | Explain techniques to limit memory usage when converting large worksheets to data URIs with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsDataUriDemo
{
    // Render the first worksheet of an Aspose.Cells workbook to a PNG image in memory, encode it as a Base64 data URI, and embed the URI directly in a CSS background‑image rule without creating a separate file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue("World");

                // Configure rendering options for PNG output (default format is PNG)
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();

                // Render the first worksheet page to a memory stream
                using (MemoryStream imageStream = new MemoryStream())
                {
                    WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);
                    // Render the first sheet directly into the stream as PNG
                    renderer.ToImage(0, imageStream);

                    // Get the PNG bytes
                    byte[] pngBytes = imageStream.ToArray();

                    // Convert to Base64 and build the data URI
                    string base64 = Convert.ToBase64String(pngBytes);
                    string dataUri = $"data:image/png;base64,{base64}";

                    // Example CSS using the data URI as a background image
                    string css = $".worksheet-background {{ background-image: url('{dataUri}'); }}";

                    // Output the CSS string
                    Console.WriteLine("Generated CSS:");
                    Console.WriteLine(css);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
