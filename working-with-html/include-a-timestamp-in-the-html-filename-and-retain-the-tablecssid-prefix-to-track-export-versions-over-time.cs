// Title: Export an Excel workbook to HTML with a timestamped filename while preserving the TableCssId prefix using Aspose.Cells for .NET
// AI Prompts: Create C# code that loads or creates a Workbook, sets HtmlSaveOptions.TableCssId to a fixed prefix, and saves the workbook as an HTML file whose name includes the current date and time. | Adapt existing Aspose.Cells HTML export logic to produce a filename such as Export_20231127_154530.html and keep the TableCssId value unchanged across multiple exports.
// Common Searches: Aspose.Cells C# export workbook to HTML with date and time in the exported file name | how to keep TableCssId prefix when saving Excel as HTML using Aspose.Cells | dynamic HTML file name generation for Aspose.Cells export in .NET | track HTML export versions by adding a timestamp to the filename with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions with date‑time filename | retain TableCssId during HTML export Aspose.Cells | C# Excel to HTML export with dynamic file name | HTML export version tracking Aspose.Cells | generate timestamped HTML output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example loads an existing Excel file or creates a new workbook, configures HtmlSaveOptions (including preserving the TableCssId prefix), builds a filename like Export_20231127_154530.html using the current timestamp, saves the workbook as HTML, and writes the output path to the console.
    class HtmlExportWithTimestamp
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Cells["A1"].PutValue("Sample Data");
                }

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                // Additional option customizations can be added here if needed

                // Generate timestamped HTML filename
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string htmlFileName = $"Export_{timestamp}.html";

                // Save the workbook as HTML
                workbook.Save(htmlFileName, htmlOptions);
                Console.WriteLine($"Workbook exported successfully to {htmlFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
