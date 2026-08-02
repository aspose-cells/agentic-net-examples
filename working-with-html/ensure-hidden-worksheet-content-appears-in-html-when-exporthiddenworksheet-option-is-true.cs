// Title: Export Hidden Worksheets to HTML with Aspose.Cells for .NET
// Description: Demonstrates how to set HtmlSaveOptions.ExportHiddenWorksheet = true (and ExportActiveWorksheetOnly = false) so that hidden sheets are rendered in the generated HTML file.
// Keywords: Aspose.Cells | .NET | C# | HtmlSaveOptions | ExportHiddenWorksheet | hidden worksheet HTML export | save workbook as HTML | include hidden sheets | Aspose.Cells HTML output | ExportActiveWorksheetOnly
// Common Searches: Aspose.Cells export hidden sheet to HTML C# | HtmlSaveOptions ExportHiddenWorksheet example | How to include hidden worksheets in HTML output | Export entire workbook with hidden sheets Aspose.Cells | C# generate HTML from workbook with hidden sheets
// Developer Intent: Create an HTML file that contains data from both visible and hidden worksheets.
// Use Cases: Produce a web‑ready report that shows supplemental data stored in hidden tabs. | Archive a workbook as HTML while preserving all sheet content for compliance audits. | Provide clients with a full HTML view of a multi‑sheet workbook, including hidden calculations.
// AI Prompts: Show how to export only hidden worksheets to HTML using Aspose.Cells. | Give a C# example that customizes the HTML style while including hidden sheets. | Explain the interaction between ExportHiddenWorksheet and ExportActiveWorksheetOnly. | Suggest ways to embed images and CSS when exporting hidden worksheets to HTML.

using System;
using Aspose.Cells;

namespace AsposeCellsHiddenWorksheetHtmlDemo
{
    // Demonstrates how to set HtmlSaveOptions.ExportHiddenWorksheet = true (and ExportActiveWorksheetOnly = false) so that hidden sheets are rendered in the generated HTML file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // First worksheet – visible
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Data from visible sheet");

            // Second worksheet – hidden
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Data from hidden sheet");
            hiddenSheet.IsVisible = false; // Mark the sheet as hidden

            // Configure HTML save options to export hidden worksheets
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = true,   // Ensure hidden sheet content is included
                ExportActiveWorksheetOnly = false // Export the whole workbook
            };

            // Save the workbook as HTML; hidden sheet data will appear in the output
            workbook.Save("WorkbookWithHiddenSheet.html", htmlOptions);

            Console.WriteLine("HTML file generated with hidden worksheet content included.");
        }
    }
}
