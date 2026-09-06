// Title: Export an Excel workbook to a single UTF-8 encoded HTML file with grid lines using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file and saves it as a UTF-8 HTML document with grid lines using Aspose.Cells. | Demonstrate how to set HtmlSaveOptions.Encoding to UTF-8 and export the full workbook as one HTML page. | Show a complete example that configures HtmlSaveOptions to include grid lines and produces a single UTF-8 HTML output.
// Common Searches: how to save an Excel workbook as UTF-8 HTML with Aspose.Cells in C# | Aspose.Cells export full workbook to one HTML file including grid lines | set HTML output encoding to UTF-8 using HtmlSaveOptions Aspose.Cells | C# generate single-page HTML from .xlsx with international characters | export Excel to HTML preserving Unicode characters Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions UTF-8 | C# export Excel to single HTML page | include grid lines in HTML export Aspose.Cells | full workbook HTML conversion Aspose.Cells | Unicode support HTML output Aspose.Cells

using System;
using System.Text;
using Aspose.Cells;

namespace ExcelToHtmlExport
{
    // Loads an .xlsx workbook, configures HtmlSaveOptions with Encoding = Encoding.UTF8, enables ExportGridLines and full‑workbook export, then saves the result as a single UTF-8 HTML file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the source Excel workbook
            // Replace "input.xlsx" with the path to your Excel file
            Workbook workbook = new Workbook("input.xlsx");

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Set the output encoding to UTF-8 for proper international character support
                Encoding = Encoding.UTF8,

                // Optional: Export the entire workbook as a single HTML file
                ExportActiveWorksheetOnly = false,

                // Optional: Export grid lines and other formatting as needed
                ExportGridLines = true
            };

            // Save the workbook as an HTML file with the specified encoding
            // Replace "output.html" with the desired output path
            workbook.Save("output.html", saveOptions);
        }
    }
}
