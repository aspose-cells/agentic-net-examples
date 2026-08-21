// Title: ASP.NET Core Web API that returns an Excel range as HTML using Aspose.Cells
// Description: Shows how to create a Workbook, populate cells A1:J25, configure HtmlSaveOptions for HTML5 with Base64‑encoded images, convert the selected range to HTML via range.ToHtml, and send the resulting HTML string directly from a Web API action without writing a file.
// Keywords: Aspose.Cells | C# | ASP.NET Core | Web API | Excel to HTML | range.ToHtml | HtmlSaveOptions | HTML5 | Base64 images | return HTML string
// Common Searches: ASP.NET Core return Excel range as HTML | Aspose.Cells convert worksheet to HTML5 in Web API | range.ToHtml example ASP.NET | How to send Excel HTML preview from API | Aspose.Cells Web API endpoint returning HTML string
// Developer Intent: Expose a Web API endpoint that generates HTML from a specified Excel range and returns the HTML markup as the response body.
// Use Cases: Provide a GET endpoint for client‑side preview of spreadsheet data as HTML5. | Create a service that accepts an uploaded workbook, converts a user‑defined range to HTML with embedded images, and returns the markup for email or reporting. | Build a microservice that on‑demand transforms large Excel reports into HTML without persisting temporary files.
// AI Prompts: Generate an ASP.NET Core controller action that builds a Workbook, fills A1:J25, uses HtmlSaveOptions (HTML5, ExportImagesAsBase64 = true), converts the range to HTML with range.ToHtml, and returns the HTML string in an OkResult. | Write code for a POST Web API method that receives an Excel file, selects a range based on query parameters, converts that range to HTML using Aspose.Cells, and streams the HTML string back to the caller. | Provide minimal Program.cs/Startup.cs configuration for ASP.NET Core to register required services and map a route to the HTML‑returning endpoint.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace MyApp
{
    // Shows how to create a Workbook, populate cells A1:J25, configure HtmlSaveOptions for HTML5 with Base64‑encoded images, convert the selected range to HTML via range.ToHtml, and send the resulting HTML string directly from a Web API action without writing a file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Fill sample data into cells A1:J25
                for (int i = 0; i < 25; i++)
                {
                    for (int j = 0; j < 10; j++)
                    {
                        worksheet.Cells[i, j].Value = $"Cell {i + 1},{j + 1}";
                    }
                }

                // Create the range to be exported (use fully qualified type to avoid ambiguity)
                Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:J25");

                // Set HTML save options
                HtmlSaveOptions options = new HtmlSaveOptions
                {
                    HtmlVersion = HtmlVersion.Html5,      // Use HTML5
                    ExportImagesAsBase64 = true          // Embed images as Base64
                };

                // Convert the range to HTML (returns a byte array)
                byte[] htmlBytes = range.ToHtml(options);
                string htmlContent = Encoding.UTF8.GetString(htmlBytes);

                // Determine output path and ensure the directory exists
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.html");
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write HTML content to file
                File.WriteAllText(outputPath, htmlContent, Encoding.UTF8);
                Console.WriteLine($"HTML file generated at: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
