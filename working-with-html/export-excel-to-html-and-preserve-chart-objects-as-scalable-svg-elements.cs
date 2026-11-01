// Title: How to convert an .xlsx workbook to HTML with embedded chart images using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel workbook to HTML with charts embedded as Base64 | Save Excel file as HTML preserving chart graphics in Aspose.Cells .NET | Convert .xlsx to HTML and embed chart images using HtmlSaveOptions in C# | C# Aspose.Cells generate HTML output with inline chart images | How to embed Excel chart objects in HTML using Aspose.Cells for .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | C# export Excel to HTML with embedded charts | Aspose.Cells convert .xlsx to HTML with inline images | HTML output base64 chart images Aspose.Cells | Aspose.Cells chart rendering in HTML

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // Loads an Excel workbook, configures HtmlSaveOptions to embed all images—including charts—as Base64 strings, and saves the result as an HTML file.
    class ExcelToHtmlWithSvgCharts
    {
        static void Main()
        {
            try
            {
                // Path to the source Excel file
                string excelPath = "input.xlsx";

                // Path where the resulting HTML file will be saved
                string htmlPath = "output.html";

                // Verify that the input file exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Input file not found: {excelPath}");
                    return;
                }

                // Load the workbook from the Excel file
                Workbook workbook = new Workbook(excelPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Embed images (including charts) directly into the HTML as Base64 strings
                    ExportImagesAsBase64 = true
                };

                // Save the workbook as an HTML file with the specified options
                workbook.Save(htmlPath, htmlOptions);
                Console.WriteLine($"HTML file saved successfully to: {htmlPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
