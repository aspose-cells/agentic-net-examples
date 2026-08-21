// Title: C# – Export Excel to Responsive HTML without Hidden Sheets using Aspose.Cells (ExportHiddenWorksheet = false, WidthScalable = true)
// Description: Demonstrates how to create a workbook with a visible and a hidden worksheet, then save it as HTML while omitting hidden sheets (ExportHiddenWorksheet = false) and enabling column‑width scaling (WidthScalable = true) for a responsive layout.
// Keywords: Aspose.Cells HTML export | ExportHiddenWorksheet false | WidthScalable true | responsive HTML from Excel | C# Aspose.Cells example | exclude hidden worksheets HTML | scalable column widths
// Common Searches: Aspose.Cells export HTML without hidden worksheets | How to make HTML columns responsive with Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet example C# | WidthScalable option Aspose.Cells HTML export | C# code to hide sheets when saving Excel as HTML
// Developer Intent: Generate an HTML file from an Excel workbook that excludes any hidden worksheets and automatically adjusts column widths for different screen sizes.
// Use Cases: Publishing web‑ready reports that hide internal worksheets while keeping tables fluid on mobile devices. | Embedding Excel data in portals where column widths must adapt to varying container widths. | Creating public HTML exports that protect confidential data stored in hidden sheets.
// AI Prompts: Show C# code to export an Aspose.Cells workbook to HTML with ExportHiddenWorksheet set to false and WidthScalable enabled. | Explain how ExportHiddenWorksheet and WidthScalable affect the HTML output in Aspose.Cells. | Give a step‑by‑step guide for creating responsive HTML from Excel while omitting hidden worksheets using Aspose.Cells .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to create a workbook with a visible and a hidden worksheet, then save it as HTML while omitting hidden sheets (ExportHiddenWorksheet = false) and enabling column‑width scaling (WidthScalable = true) for a responsive layout.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "VisibleSheet";

            // Add some sample data
            sheet.Cells["A1"].PutValue("Header 1");
            sheet.Cells["B1"].PutValue("Header 2");
            sheet.Cells["A2"].PutValue("Data 1");
            sheet.Cells["B2"].PutValue("Data 2");

            // Add a hidden worksheet to demonstrate ExportHiddenWorksheet = false
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Hidden Data");
            hiddenSheet.IsVisible = false;

            // Configure HTML save options:
            // - Do not export hidden worksheets
            // - Use scalable column widths for responsive layout
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = false,
                WidthScalable = true
            };

            // Save the workbook as HTML with the specified options
            workbook.Save("output_responsive.html", saveOptions);
        }
    }
}
