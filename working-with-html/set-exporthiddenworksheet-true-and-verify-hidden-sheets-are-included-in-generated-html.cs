// Title: Include Hidden Worksheets in HTML Export with Aspose.Cells for .NET
// Description: Demonstrates how to set HtmlSaveOptions.ExportHiddenWorksheet to true, export the entire workbook (including hidden sheets) to a single HTML file, and programmatically verify that hidden‑sheet data appears in the generated markup.
// Keywords: Aspose.Cells HTML export hidden sheet | ExportHiddenWorksheet true | C# Aspose.Cells hidden worksheet | verify hidden sheet in HTML | ExportActiveWorksheetOnly false
// Common Searches: Aspose.Cells export hidden worksheets to HTML | HtmlSaveOptions ExportHiddenWorksheet example C# | how to include hidden sheets in HTML output Aspose.Cells | check hidden sheet content after HTML conversion
// Developer Intent: Export a workbook to HTML while preserving hidden worksheets and confirm their content is present in the output file.
// Use Cases: Create an HTML audit report that shows data from both visible and hidden tabs. | Publish a web‑ready version of a workbook where supplemental information resides on hidden sheets. | Automate testing to ensure hidden‑sheet data is not omitted during HTML conversion.
// AI Prompts: Show C# code that enables ExportHiddenWorksheet in Aspose.Cells and validates hidden sheet data in the resulting HTML. | Provide a snippet to export an entire workbook, including hidden worksheets, to a single HTML file using Aspose.Cells for .NET. | Explain the interaction between ExportHiddenWorksheet and ExportActiveWorksheetOnly when saving a workbook as HTML.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHiddenSheetHtmlDemo
{
    // Demonstrates how to set HtmlSaveOptions.ExportHiddenWorksheet to true, export the entire workbook (including hidden sheets) to a single HTML file, and programmatically verify that hidden‑sheet data appears in the generated markup.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add data to the default (visible) sheet
            Workbook workbook = new Workbook();
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Data in visible sheet");

            // Add a hidden worksheet and put some data in it
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Data in hidden sheet");
            hiddenSheet.IsVisible = false; // Mark the sheet as hidden

            // Configure HTML save options to export hidden worksheets
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = true,          // Ensure hidden sheets are exported
                ExportActiveWorksheetOnly = false      // Export the whole workbook
            };

            // Define output HTML file path
            string outputHtmlPath = Path.Combine(Environment.CurrentDirectory, "WorkbookWithHiddenSheet.html");

            // Save the workbook as HTML
            workbook.Save(outputHtmlPath, htmlOptions);

            // Simple verification: read the generated HTML and check for hidden sheet data
            string htmlContent = File.ReadAllText(outputHtmlPath);
            bool containsHiddenData = htmlContent.Contains("Data in hidden sheet");

            Console.WriteLine($"HTML file saved to: {outputHtmlPath}");
            Console.WriteLine($"Hidden sheet data present in HTML: {containsHiddenData}");
        }
    }
}
