// Title: Implement an ASP.NET Core Web API action that returns Excel-to-HTML conversion as a string using Aspose.Cells
// AI Prompts: Write an ASP.NET Core controller method that loads an uploaded Excel file with Aspose.Cells, configures HtmlSaveOptions to export only the first worksheet and embed images as Base64, then returns the generated HTML string in the HTTP response. | Refactor the console example into a Web API endpoint that accepts a file path parameter, converts the workbook to HTML via a MemoryStream, and sends the HTML content back without creating a physical file. | Demonstrate how to set HtmlSaveOptions (ExportActiveWorksheetOnly, ExportImagesAsBase64) and use Workbook.Save to a MemoryStream for returning HTML from a .NET Web API.
// Common Searches: asp.net core web api convert excel file to html string using aspose.cells | return html from excel workbook in asp.net core controller | aspose.cells htmlsaveoptions export active worksheet only base64 images | memorystream excel to html conversion asp.net core endpoint | how to send excel to html conversion result as response body in .net api
// Tags: aspnet core aspose.cells html conversion endpoint | export active worksheet to html memory stream | embed excel images as base64 in html output | return html string from workbook in web api | htmlsaveoptions exportactiveworksheetonly aspose.cells

using System;
using System.IO;
using Aspose.Cells;

namespace MyExcelApp
{
    // Loads an Excel workbook, configures HtmlSaveOptions to export only the first worksheet and embed images as Base64, saves to a MemoryStream, reads the stream into a string, and returns the HTML string from a Web API action.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect the first argument to be the Excel file path.
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: MyExcelApp <excelFilePath> [outputHtmlPath]");
                return;
            }

            string excelPath = args[0];

            // Verify the Excel file exists.
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: File not found - {excelPath}");
                return;
            }

            // Determine output HTML path.
            string outputPath = args.Length > 1 ? args[1] : Path.ChangeExtension(excelPath, ".html");

            try
            {
                // Load the workbook.
                var workbook = new Workbook(excelPath);

                // Configure HTML save options.
                var htmlOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = true,   // Export only the first worksheet.
                    ExportImagesAsBase64 = true         // Embed images directly.
                    // Default chart image format is PNG; other options are not required.
                };

                // Save the workbook to a memory stream in HTML format.
                string htmlContent;
                using (var memoryStream = new MemoryStream())
                {
                    workbook.Save(memoryStream, htmlOptions);
                    memoryStream.Position = 0;
                    using (var reader = new StreamReader(memoryStream))
                    {
                        htmlContent = reader.ReadToEnd();
                    }
                }

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write the HTML content to the output file.
                File.WriteAllText(outputPath, htmlContent);
                Console.WriteLine($"HTML saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during processing.
                Console.WriteLine($"Error processing file: {ex.Message}");
            }
        }
    }
}
