// Title: Export Visible Worksheets to HTML with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook with one visible and one hidden worksheet, sets HtmlSaveOptions.ExportHiddenWorksheet = false (and ExportActiveWorksheetOnly = false), and saves the workbook as an HTML file. The resulting HTML contains only the visible worksheet.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | ExportHiddenWorksheet | ExportActiveWorksheetOnly | visible worksheets HTML | hide worksheet export | convert workbook to HTML | exclude hidden sheets | Aspose.Cells HTML export
// Common Searches: Aspose.Cells export visible sheets to HTML | HtmlSaveOptions ExportHiddenWorksheet false C# | Save workbook as HTML without hidden worksheets | How to hide worksheets from HTML export Aspose.Cells | Export only visible worksheets using Aspose.Cells .NET
// Developer Intent: Export a workbook to HTML while omitting any hidden worksheets.
// Use Cases: Generate an HTML report that shows only user‑visible data, excluding confidential or auxiliary sheets. | Create a web‑ready view of a workbook where hidden tabs must not be published. | Automate batch conversion of workbooks to HTML for publishing, ensuring hidden worksheets are filtered out.
// AI Prompts: Show me how to modify the code to also exclude charts from hidden worksheets when exporting to HTML with Aspose.Cells. | Provide a C# example that saves each visible worksheet to a separate HTML file using Aspose.Cells. | Explain how ExportActiveWorksheetOnly interacts with ExportHiddenWorksheet in HtmlSaveOptions.

using System;
using Aspose.Cells;

namespace ExportVisibleWorksheetsDemo
{
    // This C# example creates a workbook with one visible and one hidden worksheet, sets HtmlSaveOptions.ExportHiddenWorksheet = false (and ExportActiveWorksheetOnly = false), and saves the workbook as an HTML file. The resulting HTML contains only the visible worksheet.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with default worksheet
            Workbook workbook = new Workbook();

            // Access the first (default) worksheet and add some data
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Data in visible sheet");

            // Add a second worksheet and hide it
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Data in hidden sheet");
            hiddenSheet.IsVisible = false; // Mark the sheet as hidden

            // Configure HTML save options to exclude hidden worksheets
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = false, // Do not export hidden worksheets
                ExportActiveWorksheetOnly = false // Export the whole workbook (visible sheets only)
            };

            // Save the workbook to HTML; only the visible worksheet will be included
            workbook.Save("VisibleSheetsOnly.html", saveOptions);
        }
    }
}
