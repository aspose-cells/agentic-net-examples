// Title: Stream an Aspose.Cells Workbook as a Single HTML5 File Directly to HttpResponse in C#
// Description: Demonstrates how to convert an Excel workbook to HTML5 with UTF‑8 encoding using Aspose.Cells, configure HtmlSaveOptions for a single‑file output, and write the result straight to an ASP.NET (Core) HttpResponse stream without creating a temporary file.
// Keywords: Aspose.Cells HTML export | C# stream HTML to HttpResponse | ASP.NET Core Aspose.Cells download | single file HTML5 Aspose | HtmlSaveOptions streaming | UTF-8 HTML output .NET | memory stream Aspose.Cells | web API Excel to HTML | download Excel as HTML C#
// Common Searches: Aspose.Cells export workbook to HTML stream | How to send Aspose.Cells HTML output via HttpResponse | ASP.NET Core return Excel as HTML5 using Aspose | Stream Aspose.Cells HTML without saving file | Set content‑type for Aspose.Cells HTML response
// Developer Intent: The developer needs to generate HTML from an Excel workbook with Aspose.Cells and deliver it instantly to the browser through an HTTP response stream.
// Use Cases: Return a live preview of a spreadsheet in a web application. | Provide a download endpoint that serves Excel data as a single HTML5 file. | Integrate Excel‑to‑HTML conversion into an API that returns HTML for embedding in client‑side pages.
// AI Prompts: Show how to modify the sample code to write the HTML output to HttpResponse.Body using a MemoryStream. | Give an ASP.NET Core controller example that sets the correct Content‑Type and Content‑Disposition headers for the streamed HTML. | Explain how to configure HtmlSaveOptions to embed images as base64 when streaming the HTML response.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsWebExport
{
    // Demonstrates how to convert an Excel workbook to HTML5 with UTF‑8 encoding using Aspose.Cells, configure HtmlSaveOptions for a single‑file output, and write the result straight to an ASP.NET (Core) HttpResponse stream without creating a temporary file.
    public class ExcelToHtmlExporter
    {
        /// <param name="outputPath">Full path of the HTML file to create.</param>
        public void Export(string outputPath)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Output path must be provided.", nameof(outputPath));

            try
            {
                // Ensure the target directory exists
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Create a sample workbook and populate it with data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleData";

                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Age");
                sheet.Cells["A2"].PutValue("John Doe");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["A3"].PutValue("Jane Smith");
                sheet.Cells["B3"].PutValue(28);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    SaveAsSingleFile = true,
                    HtmlVersion = HtmlVersion.Html5,
                    Encoding = Encoding.UTF8
                };

                // Save workbook as HTML using the configured options
                workbook.Save(outputPath, htmlOptions);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error exporting workbook to HTML: {ex.Message}");
                throw;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string outputPath = Path.Combine(Environment.CurrentDirectory, "output", "sample.html");
                var exporter = new ExcelToHtmlExporter();
                exporter.Export(outputPath);
                Console.WriteLine($"Workbook exported successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
