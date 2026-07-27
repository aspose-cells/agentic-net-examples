// Title: Export Excel to Responsive HTML (exclude hidden sheets) – Aspose.Cells .NET
// Description: Demonstrates how to save a workbook as HTML with hidden worksheets omitted (ExportHiddenWorksheet = false) and column widths set to scale responsively (WidthScalable = true) using Aspose.Cells for C#.
// Keywords: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet false | WidthScalable responsive HTML Aspose.Cells | C# export Excel to HTML without hidden sheets | Aspose.Cells responsive column widths | HTML export hidden worksheet exclusion
// Common Searches: Aspose.Cells export hidden worksheets false C# | How to make HTML columns responsive with Aspose.Cells | Save Excel as HTML without hidden sheets Aspose | WidthScalable true example Aspose.Cells | Responsive HTML output from Excel using Aspose.Cells
// Developer Intent: Generate an HTML file from a workbook that hides any hidden worksheets and uses scalable column widths for a mobile‑friendly layout.
// Use Cases: Create web reports that display only visible data while keeping configuration sheets private. | Build responsive HTML tables that adapt to different screen sizes without manual CSS. | Produce dashboard components where hidden worksheets must not be exposed to end users.
// AI Prompts: Show C# code to embed images inline while keeping ExportHiddenWorksheet false. | Generate separate responsive HTML files for each visible worksheet using Aspose.Cells. | Explain how WidthScalable converts Excel column widths to percentage‑based HTML styles.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to save a workbook as HTML with hidden worksheets omitted (ExportHiddenWorksheet = false) and column widths set to scale responsively (WidthScalable = true) using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Visible Data");

            // Add a hidden worksheet
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Hidden Data");
            hiddenSheet.IsVisible = false; // Mark worksheet as hidden

            // Configure HTML save options:
            // - Do not export hidden worksheets
            // - Use scalable column widths for responsive layout
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = false,
                WidthScalable = true
            };

            // Save the workbook to HTML using the configured options
            workbook.Save("output_responsive.html", saveOptions);
        }
    }
}
