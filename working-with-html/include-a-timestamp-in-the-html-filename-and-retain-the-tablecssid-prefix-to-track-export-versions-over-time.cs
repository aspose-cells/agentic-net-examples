// Title: Export an Excel workbook to HTML with a timestamped filename while preserving the default TableCssId prefix using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads or creates a workbook and saves it as an HTML file whose name contains the current date and time in yyyyMMdd_HHmmss format. | Show how to use Aspose.Cells HtmlSaveOptions to export a workbook to HTML while keeping the default TableCssId prefix. | Provide a try‑catch example that creates a timestamped HTML file name, applies HtmlSaveOptions, and calls Workbook.Save.
// Common Searches: Aspose.Cells .NET export workbook to HTML with date‑time stamp in file name | How to keep the default TableCssId prefix when saving Excel as HTML using Aspose.Cells | C# generate unique HTML file name for Excel to HTML conversion with Aspose.Cells
// Tags: Aspose.Cells HTML export with datetime file naming | TableCssId handling in HTML output Aspose.Cells | C# workbook to HTML conversion using HtmlSaveOptions | custom HTML filename pattern for Excel export Aspose.Cells | exception handling for Workbook.Save HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing 'input.xlsx' workbook or creates a new one with sample data, configures HtmlSaveOptions, builds a filename like 'ExportTable_20231127_154530.html' using the current date and time, and saves the workbook as HTML while retaining the default TableCssId prefix, all wrapped in a try‑catch block to handle errors.
class Program
{
    static void Main()
    {
        try
        {
            // Define input workbook path
            string inputPath = "input.xlsx";

            // Load workbook if file exists; otherwise create a new workbook with sample data
            Workbook workbook;
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
            // Note: TableCssIdPrefix is not available in this version of Aspose.Cells for .NET

            // Create a timestamped filename
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"ExportTable_{timestamp}.html";

            // Save the workbook as HTML
            workbook.Save(fileName, htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
