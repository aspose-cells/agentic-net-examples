// Title: ASP.NET Core Web API endpoint that converts an Excel workbook to an HTML string using Aspose.Cells
// Common Searches: how to convert an Excel file to HTML string in ASP.NET Core Web API with Aspose.Cells | Aspose.Cells HTMLSaveOptions base64 images example for .NET Web API | return workbook as HTML content from a C# Web API endpoint without saving a file | using MemoryStream to get HTML from Excel in ASP.NET Core controller | export only active worksheet to HTML in Aspose.Cells for a Web API response
// Tags: Aspose.Cells HTML conversion to string | HTMLSaveOptions embed images base64 | export active worksheet as HTML Aspose.Cells | ASP.NET Core controller returning HTML content | MemoryStream Excel to HTML C# | C# Web API endpoint without file output | Aspose.Cells workbook to HTML response

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example shows how to load an Excel workbook with Aspose.Cells, configure HtmlSaveOptions to embed images as Base64 and export only the active worksheet, save the workbook to a MemoryStream, read the stream as a UTF-8 HTML string, and return that string from an ASP.NET Core Web API endpoint, eliminating the need for a temporary file.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Expect at least the Excel file path; optional second argument for output HTML file.
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ExcelHtmlConverter <excelFilePath> [outputHtmlPath]");
                return;
            }

            string excelPath = args[0];

            // Verify the Excel file exists to avoid FileNotFoundException.
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: File not found - {excelPath}");
                return;
            }

            string htmlPath = args.Length > 1 ? args[1] : null;

            // Load the workbook using Aspose.Cells.
            Workbook workbook = new Workbook(excelPath);

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true,          // Embed images as Base64.
                ExportActiveWorksheetOnly = true     // Export only the active worksheet.
            };

            // Save the workbook to a memory stream in HTML format.
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, htmlOptions);
                htmlStream.Position = 0; // Reset stream position for reading.

                // Read the HTML content as a UTF-8 string.
                string htmlContent;
                using (StreamReader reader = new StreamReader(htmlStream, Encoding.UTF8))
                {
                    htmlContent = reader.ReadToEnd();
                }

                // Output the HTML either to a file or to the console.
                if (!string.IsNullOrEmpty(htmlPath))
                {
                    File.WriteAllText(htmlPath, htmlContent, Encoding.UTF8);
                    Console.WriteLine($"HTML saved to {htmlPath}");
                }
                else
                {
                    Console.WriteLine(htmlContent);
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message.
            Console.WriteLine($"Exception: {ex.Message}");
        }
    }
}
