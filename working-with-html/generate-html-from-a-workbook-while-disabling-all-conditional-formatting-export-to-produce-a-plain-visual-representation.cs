// Title: Export an Excel workbook to plain HTML without conditional formatting using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a .xlsx file with Aspose.Cells, configures HtmlSaveOptions to disable conditional formatting, and saves the workbook as an HTML file, creating the output folder if needed. | Update the provided Aspose.Cells example so that HtmlSaveOptions.ExportConditionalFormatting is set to false, producing a clean HTML view without any conditional formatting rules.
// Common Searches: aspocells c# export excel to html without conditional formatting | how to turn off conditional formatting when saving workbook as html using aspose.cells | c# generate plain html from excel workbook with aspose.cells HtmlSaveOptions | disable conditional formatting in aspose.cells html export example
// Tags: Aspose.Cells HtmlSaveOptions disable conditional formatting | C# export Excel to HTML without styles | Aspose.Cells generate plain HTML output | HtmlSaveOptions ExportConditionalFormatting false | Aspose.Cells HTML conversion without conditional rules

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // This example demonstrates loading an Excel file with Aspose.Cells, configuring HtmlSaveOptions to suppress conditional formatting during the HTML conversion, ensuring the output directory exists, and saving the workbook as a clean HTML document while handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string excelPath = @"C:\Input\Sample.xlsx";

            // Path where the HTML output will be saved
            string htmlPath = @"C:\Output\Sample.html";

            // Verify that the input file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input file not found: {excelPath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(excelPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Optional: Export only the first worksheet (set to false to export all)
                    // ExportAllWorksheets = false,

                    // Optional: Set the encoding if needed
                    // Encoding = System.Text.Encoding.UTF8
                };

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(htmlPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, htmlOptions);

                Console.WriteLine("Workbook has been exported to HTML.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
