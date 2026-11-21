// Title: Export an Excel workbook to HTML with a timestamped filename while retaining the TableCssId prefix using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export Excel to HTML with timestamp in file name | how to keep TableCssId prefix when saving workbook as HTML using Aspose.Cells | generate versioned HTML file name from Excel workbook in .NET | Aspose.Cells HtmlSaveOptions custom filename pattern with date and time | save Excel as HTML with dynamic naming and table CSS ID in C#
// Tags: Aspose.Cells HtmlSaveOptions custom filename | timestamped HTML export from Excel C# | preserve TableCssId during HTML conversion | versioned HTML file naming Aspose.Cells | Excel to HTML export with dynamic file name .NET

using Aspose.Cells;
using System;

// Loads an Excel workbook, sets HtmlSaveOptions.TableCssId, creates a date‑time string, builds a filename that includes the TableCssId prefix and timestamp, and saves the workbook as an HTML file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        // Preserve the TableCssId prefix for version tracking
        htmlOptions.TableCssId = "ExportTable";

        // Create a timestamp string (e.g., 20231117_153045)
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

        // Build the output HTML filename with the TableCssId prefix and timestamp
        string outputFileName = $"Export_{htmlOptions.TableCssId}_{timestamp}.html";

        // Save the workbook as HTML using the configured options
        workbook.Save(outputFileName, htmlOptions);
    }
}
