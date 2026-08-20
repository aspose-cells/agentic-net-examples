// Title: Export HTML without Hidden Worksheets using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook with visible and hidden sheets, configure HtmlSaveOptions with ExportHiddenWorksheet = false (and ExportActiveWorksheetOnly = false), save the workbook as HTML, and programmatically verify that hidden worksheets are omitted from the generated file.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportHiddenWorksheet false | C# HTML export workbook | exclude hidden sheets | verify hidden worksheet omitted | Aspose.Cells .NET example | save workbook as HTML
// Common Searches: Aspose.Cells hide hidden worksheets in HTML export | ExportHiddenWorksheet C# example | HTML export visible sheets only Aspose.Cells | Check hidden sheet presence in exported HTML | Aspose.Cells HtmlSaveOptions settings
// Developer Intent: Configure HtmlSaveOptions to skip hidden worksheets during HTML conversion and confirm that the hidden sheet does not appear in the output.
// Use Cases: Generate public HTML reports that contain only visible data while keeping confidential tabs hidden. | Automate quality checks for HTML exports by scanning the output for hidden worksheet identifiers. | Batch‑convert multiple workbooks to HTML, guaranteeing that all hidden worksheets are excluded from each file.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to HTML, exclude hidden worksheets, and validate the result by searching the HTML content. | Explain the effect of ExportHiddenWorksheet and ExportActiveWorksheetOnly on the HTML output of a workbook in Aspose.Cells. | Show how to modify the example to export only the active worksheet while still ignoring any hidden sheets.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHiddenSheetExportDemo
{
    // Demonstrates how to create a workbook with visible and hidden sheets, configure HtmlSaveOptions with ExportHiddenWorksheet = false (and ExportActiveWorksheetOnly = false), save the workbook as HTML, and programmatically verify that hidden worksheets are omitted from the generated file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ----- Visible worksheet -----
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Data in visible sheet");

            // ----- Hidden worksheet -----
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Data in hidden sheet");
            hiddenSheet.IsVisible = false; // Mark the sheet as hidden

            // Configure HTML save options to exclude hidden worksheets
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = false, // Do not export hidden worksheets
                ExportActiveWorksheetOnly = false // Export the whole workbook (visible sheets only)
            };

            // Path for the generated HTML file
            string htmlPath = "WorkbookWithoutHiddenSheet.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlPath, saveOptions);

            // Verify that the hidden sheet is omitted from the HTML output
            string htmlContent = File.ReadAllText(htmlPath);

            bool hiddenSheetPresent = htmlContent.Contains("HiddenSheet", StringComparison.OrdinalIgnoreCase);
            Console.WriteLine($"Hidden sheet name found in HTML: {hiddenSheetPresent}");
            Console.WriteLine(hiddenSheetPresent
                ? "Verification failed: hidden sheet was exported."
                : "Verification succeeded: hidden sheet was omitted.");
        }
    }
}
