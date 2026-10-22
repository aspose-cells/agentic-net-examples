// Title: Create an ASP.NET Core Web API endpoint that uses Aspose.Cells to convert an Excel file to an HTML string and returns it as the response
// AI Prompts: Write a C# ASP.NET Core controller action that receives an uploaded Excel file or a file path, invokes a helper method to convert it to HTML with Aspose.Cells, and returns the HTML string with the Content-Type set to text/html. | Generate code that adds robust error handling to the API endpoint so that a missing or invalid file returns a 400 Bad Request and any conversion exception returns a 500 Internal Server Error with a clear message. | Show how to configure HtmlSaveOptions (ExportImagesAsBase64 = true) and stream the resulting HTML directly from the controller without persisting a temporary file.
// Common Searches: asp.net core web api convert excel to html using aspose.cells | return html string from asp.net core controller after excel conversion | how to set HtmlSaveOptions ExportImagesAsBase64 in asp.net core api | c# endpoint that accepts uploaded .xlsx and returns html response | asp.net core return 400 when excel file path is missing
// Tags: Aspose.Cells Excel to HTML conversion in ASP.NET Core | Web API endpoint returning HTML string | HtmlSaveOptions ExportImagesAsBase64 usage | C# memory stream HTML output from workbook | Error handling for file not found in ASP.NET Core API | Upload Excel file and stream HTML response

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace MyApp
{
    // The example shows how to load an Excel workbook with Aspose.Cells, configure HtmlSaveOptions to embed images as Base64, save the workbook to a MemoryStream in HTML format, read the stream into a UTF‑8 string, and expose this conversion through an ASP.NET Core Web API action that returns the HTML string as the HTTP response, with proper error handling for missing files and conversion failures.
    public static class ExcelConverter
    {
        // Converts an Excel file to an HTML string.
        public static string ConvertExcelToHtml(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path must be provided.", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel file not found.", filePath);

            try
            {
                // Load the workbook from the specified file.
                var workbook = new Workbook(filePath);

                // Set HTML save options.
                var htmlOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true
                };

                // Save to a memory stream in HTML format using the options.
                using var memoryStream = new MemoryStream();
                workbook.Save(memoryStream, htmlOptions);
                memoryStream.Position = 0;

                // Read and return the HTML content.
                using var reader = new StreamReader(memoryStream, Encoding.UTF8);
                return reader.ReadToEnd();
            }
            catch (Exception ex)
            {
                // Wrap any exception for the caller.
                throw new InvalidOperationException("Failed to convert Excel to HTML.", ex);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ExcelConverter <excelFilePath>");
                return;
            }

            var filePath = args[0];
            try
            {
                string html = ExcelConverter.ConvertExcelToHtml(filePath);
                Console.WriteLine(html);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
