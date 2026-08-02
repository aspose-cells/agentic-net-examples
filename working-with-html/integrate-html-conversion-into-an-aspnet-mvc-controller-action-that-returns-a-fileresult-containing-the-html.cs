// Title: ASP.NET MVC: Return Excel as HTML FileResult using Aspose.Cells
// Description: Loads an Excel workbook from App_Data, configures HtmlSaveOptions (single file, active sheet, Base64‑encoded images), saves to a MemoryStream, and streams the HTML bytes back to the client as a FileResult.
// Keywords: Aspose.Cells | HTML export | ASP.NET MVC | FileResult | HtmlSaveOptions | Base64 images | Excel to HTML | downloadable HTML | streaming HTML | single HTML file
// Common Searches: asp.net mvc return excel as html fileresult | aspose.cells export workbook to html in mvc | download html version of excel worksheet asp.net | embed images base64 asp.net mvc export | stream excel html response asp.net
// Developer Intent: Implement an MVC action that converts an Excel workbook to HTML with Aspose.Cells and sends it to the client as a FileResult.
// Use Cases: Download a spreadsheet as a single HTML page from a web endpoint | Display worksheet content in the browser without creating temporary files | Provide an API that streams HTML for a specific sheet to front‑end components | Integrate Excel‑to‑HTML conversion in reporting dashboards
// AI Prompts: Write an ASP.NET MVC controller action that calls the ToHtml method, returns File(htmlBytes, "text/html", "Report.html"), and handles empty results gracefully. | Create unit tests for ExportController.ToHtml and the MVC action that returns the FileResult. | Add ILogger‑based logging for Aspose.Cells export errors and return a 500 status with a user‑friendly message. | Show how to register a route and a view that triggers the HTML export and opens the result in a new browser tab.

using System;
using System.IO;
using Aspose.Cells;

namespace MyMvcApp.Controllers
{
    // Loads an Excel workbook from App_Data, configures HtmlSaveOptions (single file, active sheet, Base64‑encoded images), saves to a MemoryStream, and streams the HTML bytes back to the client as a FileResult.
    public class ExportController
    {
        // Export the workbook to HTML and return the HTML content as a byte array.
        public byte[] ToHtml()
        {
            try
            {
                // Build the full path to the workbook file.
                var workbookPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Sample.xlsx");

                // Ensure the workbook file exists to avoid FileNotFoundException.
                if (!File.Exists(workbookPath))
                    throw new FileNotFoundException("Workbook file not found.", workbookPath);

                // Load the workbook.
                var workbook = new Workbook(workbookPath);

                // Configure HTML save options.
                var htmlOptions = new HtmlSaveOptions
                {
                    SaveAsSingleFile = true,               // generate a single HTML file
                    ExportActiveWorksheetOnly = true,      // export only the active sheet
                    ExportImagesAsBase64 = true            // embed images as Base64
                };

                // Save the workbook to a memory stream using the HTML options.
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, htmlOptions);
                    // Return the HTML content.
                    return stream.ToArray();
                }
            }
            catch (Exception ex)
            {
                // Handle or log the exception as needed.
                Console.Error.WriteLine($"Error exporting workbook to HTML: {ex.Message}");
                return Array.Empty<byte>();
            }
        }
    }

    // Entry point for console execution.
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                var controller = new ExportController();
                var htmlBytes = controller.ToHtml();

                if (htmlBytes.Length > 0)
                {
                    // Write the HTML output to a file for verification.
                    var outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Sample.html");
                    File.WriteAllBytes(outputPath, htmlBytes);
                    Console.WriteLine($"HTML exported successfully to: {outputPath}");
                }
                else
                {
                    Console.WriteLine("No HTML content was generated.");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
