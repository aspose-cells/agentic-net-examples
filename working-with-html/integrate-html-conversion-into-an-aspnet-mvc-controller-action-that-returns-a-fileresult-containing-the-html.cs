// Title: ASP.NET MVC FileResult – Export Excel to HTML using Aspose.Cells
// Description: Shows how to load an Excel workbook, apply HtmlSaveOptions for a single, mobile‑friendly HTML file, write the result to a MemoryStream, and stream the HTML bytes from an MVC controller as a FileResult.
// Keywords: Aspose.Cells | Excel to HTML | ASP.NET MVC | FileResult | HtmlSaveOptions | MemoryStream | C# export workbook | download HTML | controller action | streaming response
// Common Searches: Aspose.Cells export Excel to HTML MVC | ASP.NET MVC return HTML file from controller | download Excel as HTML using Aspose.Cells | FileResult streaming HTML bytes C# | convert workbook to HTML in ASP.NET MVC
// Developer Intent: Create an MVC controller action that converts an Excel file to HTML with Aspose.Cells and returns the generated content as a downloadable FileResult.
// Use Cases: Expose an ExportHtml endpoint that calls ExcelExporter.ExportToHtml() and returns File(htmlBytes, "text/html", "report.html"). | Convert an uploaded Excel workbook to HTML on the fly and embed the markup in an email while also offering a download link. | Generate HTML reports server‑side, save them to disk with SaveHtmlToFile, and serve the saved files via FilePathResult for archival access.
// AI Prompts: Write an ASP.NET MVC controller action named ExportHtml that uses the provided ExcelExporter class to convert sample.xlsx to HTML and returns a FileResult with the appropriate content type and filename. | Show how to modify HtmlSaveOptions to embed images as base64 strings and stream the resulting HTML through a FileResult in MVC. | Create a unit test for the ExportHtml action that verifies the returned FileResult contains non‑empty HTML bytes and the expected file name.

using System;
using System.IO;
using Aspose.Cells;

namespace MyMvcApp
{
    // Shows how to load an Excel workbook, apply HtmlSaveOptions for a single, mobile‑friendly HTML file, write the result to a MemoryStream, and stream the HTML bytes from an MVC controller as a FileResult.
    public class ExcelExporter
    {
        /// <returns>Byte array containing the HTML content.</returns>
        public byte[] ExportToHtml()
        {
            try
            {
                // Build the full path to the Excel file (adjust if needed)
                string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "sample.xlsx");

                // Ensure the source file exists to avoid FileNotFoundException
                if (!File.Exists(excelPath))
                {
                    throw new FileNotFoundException("The Excel file was not found.", excelPath);
                }

                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    SaveAsSingleFile = true,               // Export as a single HTML file
                    ExportActiveWorksheetOnly = true,      // Export only the active sheet (optional)
                    IsMobileCompatible = true              // Make the output mobile‑friendly (optional)
                };

                // Save the workbook to a memory stream using the HTML options
                using (MemoryStream htmlStream = new MemoryStream())
                {
                    workbook.Save(htmlStream, htmlOptions);
                    return htmlStream.ToArray(); // Return the HTML content as a byte array
                }
            }
            catch (Exception ex)
            {
                // Wrap and rethrow for higher‑level handling or logging
                throw new ApplicationException("Failed to export Excel to HTML.", ex);
            }
        }

        /// <param name="outputPath">Full path where the HTML file will be saved.</param>
        public void SaveHtmlToFile(string outputPath)
        {
            try
            {
                byte[] htmlBytes = ExportToHtml();
                File.WriteAllBytes(outputPath, htmlBytes);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Failed to save HTML to '{outputPath}'.", ex);
            }
        }
    }

    // Simple console entry point for demonstration/testing
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                // Determine output path (same folder as executable)
                string outputHtml = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sample.html");

                ExcelExporter exporter = new ExcelExporter();
                exporter.SaveHtmlToFile(outputHtml);

                Console.WriteLine($"HTML file successfully saved to: {outputHtml}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
