// Title: Export Excel to Plain HTML without Conditional Formatting using Aspose.Cells for .NET
// Description: Shows how to load an Excel workbook with Aspose.Cells, clear all conditional formatting from each worksheet, configure HtmlSaveOptions for a basic export, and save the result as plain HTML that displays only the cell values.
// Keywords: Aspose.Cells | HTML export | remove conditional formatting | .NET | C# | Workbook.Save | HtmlSaveOptions | plain HTML | Excel to HTML conversion | clear conditional formatting
// Common Searches: Aspose.Cells export Excel to HTML without conditional formatting | C# remove conditional formatting before HTML conversion | Generate plain HTML from workbook using Aspose.Cells | Disable conditional formatting in Aspose.Cells HTML output | How to save Excel as simple HTML with Aspose.Cells
// Developer Intent: Produce an HTML file from an Excel workbook that omits all conditional formatting, yielding a clean visual snapshot of the data.
// Use Cases: Create a web‑ready view of a financial report where color rules are unnecessary. | Prepare HTML email content from Excel data while stripping formatting for consistent styling. | Provide a lightweight HTML preview of raw data for dashboards that apply their own CSS.
// AI Prompts: Generate C# code that loads an Excel workbook, removes all conditional formatting, and saves it as plain HTML using Aspose.Cells. | Explain how HtmlSaveOptions affect the HTML output when exporting a workbook with Aspose.Cells. | Show a step‑by‑step guide to export Excel to HTML without conditional formatting in .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to load an Excel workbook with Aspose.Cells, clear all conditional formatting from each worksheet, configure HtmlSaveOptions for a basic export, and save the result as plain HTML that displays only the cell values.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Remove all conditional formatting from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clear the collection of conditional formattings
                sheet.ConditionalFormattings.Clear();
            }

            // Configure HTML save options (default settings are sufficient for a plain view)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export all data (including values, not formulas) – default is All
                ExportDataOptions = HtmlExportDataOptions.All
            };

            // Save the workbook as an HTML file without conditional formatting
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook has been saved to HTML at: {outputPath}");
        }
    }
}
