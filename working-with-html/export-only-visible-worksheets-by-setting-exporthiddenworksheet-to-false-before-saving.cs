// Title: Export Visible Worksheets to HTML with Aspose.Cells (C#) – Disable ExportHiddenWorksheet
// Description: Shows how to create a workbook with visible and hidden sheets, set HtmlSaveOptions.ExportHiddenWorksheet = false, and save the workbook as HTML so that only the visible worksheets are exported.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportHiddenWorksheet | C# | .NET | export visible worksheets | HTML conversion | hide worksheet | exclude hidden sheets | workbook to HTML | Aspose.Cells example
// Common Searches: Aspose.Cells export only visible sheets to HTML | HtmlSaveOptions ExportHiddenWorksheet false example | C# save workbook as HTML without hidden worksheets | How to hide worksheets from HTML export using Aspose.Cells | Export visible worksheets only Aspose.Cells .NET
// Developer Intent: Generate an HTML file that contains only the workbook's visible worksheets, omitting any hidden tabs.
// Use Cases: Publish a web‑ready report that shows only user‑visible data while keeping analysis sheets private. | Create public documentation from Excel files, automatically excluding hidden or draft worksheets. | Automate conversion of workbooks to HTML for websites, ensuring confidential tabs are not exposed.
// AI Prompts: Provide C# code that saves an Aspose.Cells workbook to HTML while skipping hidden worksheets. | Explain the difference between ExportHiddenWorksheet and ExportActiveWorksheetOnly when converting to HTML. | Give a step‑by‑step guide to configure HtmlSaveOptions for exporting only visible sheets in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace ExportVisibleWorksheetsDemo
{
    // Shows how to create a workbook with visible and hidden sheets, set HtmlSaveOptions.ExportHiddenWorksheet = false, and save the workbook as HTML so that only the visible worksheets are exported.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Data in visible sheet");

            // Add a hidden worksheet
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
