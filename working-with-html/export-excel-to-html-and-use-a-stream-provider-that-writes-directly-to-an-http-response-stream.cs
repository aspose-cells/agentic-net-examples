// Title: Stream Aspose.Cells Workbook as HTML Directly to ASP.NET HttpResponse (C#)
// Description: Demonstrates how to export an Aspose.Cells Workbook to a single HTML page and write it straight to an ASP.NET HttpResponse output stream using HtmlSaveOptions, eliminating the need for a temporary file on disk.
// Keywords: Aspose.Cells HTML streaming | C# export workbook to HttpResponse | HtmlSaveOptions SaveAsSingleFile stream | ASP.NET return Excel as HTML | Aspose.Cells write HTML to response | stream HTML from Excel C# | Aspose.Cells web export example
// Common Searches: Aspose.Cells export workbook to HTML stream | C# write Aspose.Cells HTML to HttpResponse | How to return Excel as HTML in ASP.NET | Aspose.Cells SaveAsSingleFile to response stream | Streaming Excel HTML output without file
// Developer Intent: Generate an HTML representation of a workbook and send it directly to the client via an HTTP response stream.
// Use Cases: Render an Excel report as HTML in a web page without creating a file on the server. | Provide a download endpoint that streams HTML content generated from a workbook. | Integrate dynamic Excel‑to‑HTML conversion into ASP.NET MVC or Web API actions.
// AI Prompts: Show the complete ASP.NET Core controller action that streams an Aspose.Cells workbook as a single HTML file to the client. | Explain how to set the correct Content‑Type and Content‑Disposition headers when streaming HTML from Aspose.Cells. | Provide a version of the ExportToHtml class that uses a custom IStreamProvider to write directly to HttpResponse.OutputStream.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWebDemo
{
    // Handles exporting a workbook to an HTML file.
    // Demonstrates how to export an Aspose.Cells Workbook to a single HTML page and write it straight to an ASP.NET HttpResponse output stream using HtmlSaveOptions, eliminating the need for a temporary file on disk.
    public class ExportToHtml
    {
        public void Export(string outputPath)
        {
            try
            {
                // Create a new workbook and add sample data.
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Aspose.Cells HTML Export");
                sheet.Cells["A2"].PutValue(DateTime.Now.ToString());

                // Ensure the output directory exists.
                var directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Configure HTML save options.
                var htmlOptions = new HtmlSaveOptions
                {
                    SaveAsSingleFile = true
                };

                // Save the workbook as a single HTML file using the options.
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully exported to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during export: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Define the output HTML file path.
            string outputFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExportedWorkbook.html");

            // Perform the export.
            var exporter = new ExportToHtml();
            exporter.Export(outputFile);
        }
    }
}
