// Title: C# – Export Excel to HTML with Timestamped Filename & TableCssId using Aspose.Cells
// Description: Shows how to build a yyyyMMdd_HHmmss timestamp, apply it to the HTML file name and the TableCssId prefix, and save a workbook as HTML with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# HTML export | timestamped filename | TableCssId | HtmlSaveOptions | Excel to HTML | versioned reports | dynamic file name | Aspose.Cells .NET | HTML report automation
// Common Searches: Aspose.Cells export HTML with timestamp | C# set TableCssId prefix Aspose.Cells | add date to exported HTML file name Aspose | HtmlSaveOptions timestamp example | generate versioned HTML from Excel C#
// Developer Intent: Export an Excel workbook to HTML where both the output file name and the TableCssId contain the current timestamp for unique version identification.
// Use Cases: Automated daily reporting with uniquely named HTML files for archiving. | Audit logs that map CSS selectors to specific export runs. | Client‑side scripts that target tables using a timestamped TableCssId. | Batch processing of multiple workbooks, each producing a distinct HTML identifier. | Embedding Excel data into web portals while maintaining version‑controlled styling.
// AI Prompts: Generate C# code using Aspose.Cells to save a workbook as HTML with a filename like Export_20231130_101530.html and TableCssId set to export_20231130_101530_. | Explain step‑by‑step how to configure HtmlSaveOptions for timestamped filenames and TableCssId in Aspose.Cells. | Provide a concise tutorial for creating a timestamp string, applying it to both the output file name and TableCssId, and saving the workbook as HTML. | Show how to modify the example to include a custom CSS file reference while keeping the timestamped TableCssId. | Write a PowerShell script that calls the compiled C# program to generate timestamped HTML exports for a list of Excel files.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to build a yyyyMMdd_HHmmss timestamp, apply it to the HTML file name and the TableCssId prefix, and save a workbook as HTML with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);

            // Generate a timestamp string for the file name and TableCssId
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            // Build the HTML file name with the timestamp
            string htmlFileName = $"Export_{timestamp}.html";

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Use the timestamp as a prefix for TableCssId to track export versions
            saveOptions.TableCssId = $"export_{timestamp}_";

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlFileName, saveOptions);

            Console.WriteLine($"Workbook exported to HTML file: {htmlFileName}");
            Console.WriteLine($"TableCssId prefix set to: {saveOptions.TableCssId}");
        }
    }
}
