// Title: Export Workbook to HTML with Aspose.Cells for .NET – Omit Hidden Rows and Worksheets (ExportHiddenWorksheet)
// Description: Demonstrates how to generate HTML from an Aspose.Cells workbook while excluding hidden rows and hidden worksheets unless the ExportHiddenWorksheet flag is set. The sample creates a visible sheet and a hidden sheet, hides specific rows, and saves two HTML files: one that contains only visible data, and another that also includes the hidden worksheet (still without its hidden rows).
// Keywords: Aspose.Cells HTML export | ExportHiddenWorksheet .NET | remove hidden rows Aspose.Cells | hide worksheet HTML save | HtmlSaveOptions hidden content | Aspose.Cells tutorial C# | generate HTML report Aspose
// Common Searches: Aspose.Cells export HTML without hidden rows | How to hide worksheets from HTML output in Aspose.Cells | ExportHiddenWorksheet option example | Remove hidden rows when saving workbook as HTML | C# Aspose.Cells HTMLSaveOptions hidden worksheet
// Developer Intent: Create HTML output from a workbook that shows only visible rows and sheets, with the ability to include hidden worksheets on demand.
// Use Cases: Publish a public HTML report that displays only data the user can see by setting ExportHiddenWorksheet = false and HiddenRowDisplayType = Remove. | Generate an internal HTML version that includes hidden worksheets for review while still omitting hidden rows. | Validate hidden‑content handling by comparing two HTML files—one without hidden worksheets and one with them.
// AI Prompts: Show how to export hidden rows as empty cells instead of removing them using HtmlSaveOptions. | Provide a code example that hides column headers when exporting a workbook to HTML with Aspose.Cells. | Explain how to list all hidden rows and worksheets before saving the workbook to HTML.

using System;
using Aspose.Cells;

// Demonstrates how to generate HTML from an Aspose.Cells workbook while excluding hidden rows and hidden worksheets unless the ExportHiddenWorksheet flag is set. The sample creates a visible sheet and a hidden sheet, hides specific rows, and saves two HTML files: one that contains only visible data, and another that also includes the hidden worksheet (still without its hidden rows).
class HiddenRowsAndWorksheetDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // ----- Visible worksheet -----
        Worksheet visibleSheet = workbook.Worksheets[0];
        visibleSheet.Name = "VisibleSheet";

        // Add data to visible rows
        visibleSheet.Cells["A1"].PutValue("Visible Row 1");
        visibleSheet.Cells["A2"].PutValue("Visible Row 2");
        visibleSheet.Cells["A3"].PutValue("Visible Row 3");

        // Hide the second row (index 1)
        visibleSheet.Cells.HideRow(1);

        // ----- Hidden worksheet -----
        Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");

        // Add data to the hidden sheet
        hiddenSheet.Cells["A1"].PutValue("Hidden Sheet Row 1");
        hiddenSheet.Cells["A2"].PutValue("Hidden Sheet Row 2");
        hiddenSheet.Cells["A3"].PutValue("Hidden Sheet Row 3");

        // Hide a row in the hidden sheet as well
        hiddenSheet.Cells.HideRow(1);

        // Make the entire worksheet hidden
        hiddenSheet.IsVisible = false;

        // ----- Save without exporting hidden worksheets -----
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportHiddenWorksheet = false,                 // Do not export hidden worksheets
            HiddenRowDisplayType = HtmlHiddenRowDisplayType.Remove // Remove hidden rows from HTML
        };

        workbook.Save("output_without_hidden.html", saveOptions);

        // ----- Save with exporting hidden worksheets -----
        saveOptions.ExportHiddenWorksheet = true; // Now export hidden worksheets

        workbook.Save("output_with_hidden.html", saveOptions);

        // Verification note:
        // - output_without_hidden.html will contain only the visible sheet and will not show the hidden row.
        // - output_with_hidden.html will include the hidden sheet (its visible rows only, hidden rows removed).
        Console.WriteLine("HTML files generated. Check the two output files to verify hidden rows and worksheets handling.");
    }
}
