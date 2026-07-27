// Title: ASP.NET Core Web API: Return Aspose.Cells Workbook as HTML String
// Description: Demonstrates how to create an Aspose.Cells workbook, populate it with text, numbers and dates, configure HtmlSaveOptions for HTML5, export only the active worksheet, embed images as Base64, save to a MemoryStream, read the content with the correct encoding, and return the generated HTML markup as a plain‑text response from a Web API endpoint.
// Keywords: Aspose.Cells | C# | HTML conversion | HtmlSaveOptions | ASP.NET Core | Web API | return HTML string | Excel to HTML | Base64 images | HTML5 export | active worksheet only
// Common Searches: Aspose.Cells return HTML from ASP.NET Core Web API | Convert Excel workbook to HTML string C# | HtmlSaveOptions example for Web API response | Export Aspose.Cells worksheet as HTML5 | Base64 image embedding Aspose.Cells HTML export
// Developer Intent: Provide a Web API action that converts an Aspose.Cells workbook to HTML and sends the markup back as the HTTP response body.
// Use Cases: Expose a GET endpoint that streams Excel data as instantly viewable HTML in the browser. | Integrate Excel preview into a single‑page application without creating temporary files on the server. | Deliver HTML with embedded Base64 images so the client receives a self‑contained document.
// AI Prompts: Generate an ASP.NET Core controller method that calls GenerateHtmlFromWorkbook and returns Ok(html). | Create a POST endpoint that accepts an uploaded .xlsx file, converts it to HTML using the same HtmlSaveOptions, and returns the HTML string. | Add response caching to the HTML conversion endpoint to avoid re‑processing identical workbooks.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace MyAsposeApp
{
    // Demonstrates how to create an Aspose.Cells workbook, populate it with text, numbers and dates, configure HtmlSaveOptions for HTML5, export only the active worksheet, embed images as Base64, save to a MemoryStream, read the content with the correct encoding, and return the generated HTML markup as a plain‑text response from a Web API endpoint.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string html = GenerateHtmlFromWorkbook();
                Console.WriteLine(html);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Generates HTML representation of a sample workbook.
        private static string GenerateHtmlFromWorkbook()
        {
            // Create a new workbook and add sample data.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            sheet.Cells["A2"].PutValue("This is exported as HTML.");
            sheet.Cells["B1"].PutValue(12345);
            sheet.Cells["B2"].PutValue(DateTime.Now);

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                HtmlVersion = HtmlVersion.Html5,               // Use HTML5 for modern browsers.
                ExportActiveWorksheetOnly = true,              // Export only the active worksheet.
                ExportImagesAsBase64 = true                    // Embed images as Base64.
            };

            // Save the workbook to a memory stream using the HTML options.
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, htmlOptions);
                ms.Position = 0;

                // Determine the encoding (default UTF-8 if not set).
                Encoding encoding = htmlOptions.Encoding ?? Encoding.UTF8;

                // Read and return the HTML content.
                using (StreamReader reader = new StreamReader(ms, encoding))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}
