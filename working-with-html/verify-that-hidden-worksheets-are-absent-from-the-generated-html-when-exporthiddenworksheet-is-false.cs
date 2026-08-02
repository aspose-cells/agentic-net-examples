// Title: Exclude Hidden Worksheets from HTML Export with AspNet Cells for .NET (C#)
// Description: Creates a workbook containing a visible and a hidden sheet, saves it to HTML with HtmlSaveOptions.ExportHiddenWorksheet set to false, reads the output to confirm the hidden sheet is omitted, then repeats the export with the option enabled to show the difference.
// Keywords: Aspose.Cells C# HTML export hidden worksheet | ExportHiddenWorksheet false | Aspose.Cells hide sheet in HTML | C# verify hidden sheet not exported | Aspose.Cells HtmlSaveOptions example
// Common Searches: Aspose.Cells prevent hidden sheets from appearing in HTML | C# check hidden worksheet excluded from HTML output | HtmlSaveOptions ExportHiddenWorksheet usage | Aspose.Cells hide worksheet during HTML conversion | verify hidden sheet omission Aspose.Cells
// Developer Intent: Confirm that hidden worksheets are omitted from the generated HTML when ExportHiddenWorksheet is set to false.
// Use Cases: Generate client‑facing HTML reports that show only visible worksheets. | Automate validation in CI pipelines to ensure hidden data never leaks to HTML exports. | Toggle ExportHiddenWorksheet to compare HTML outputs for debugging or documentation.
// AI Prompts: Write a C# unit test using Aspose.Cells that asserts hidden worksheets are excluded when HtmlSaveOptions.ExportHiddenWorksheet is false. | Provide a C# snippet that logs a warning if any hidden worksheet name is found in the saved HTML file. | Generate C# code that creates two HTML files—one without hidden worksheets and one with them—and prints verification results for each.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHiddenWorksheetVerification
{
    // Creates a workbook containing a visible and a hidden sheet, saves it to HTML with HtmlSaveOptions.ExportHiddenWorksheet set to false, reads the output to confirm the hidden sheet is omitted, then repeats the export with the option enabled to show the difference.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ----- Visible worksheet -----
            Worksheet visibleSheet = workbook.Worksheets[0]; // default first sheet
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Visible Data");

            // ----- Hidden worksheet -----
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Hidden Data");
            hiddenSheet.IsVisible = false; // hide the sheet

            // Configure HTML save options to NOT export hidden worksheets
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = false, // crucial setting
                ExportActiveWorksheetOnly = false // export whole workbook (visible sheets only)
            };

            // Path for the generated HTML file
            string htmlPath = "output_without_hidden.html";

            // Save the workbook as HTML
            workbook.Save(htmlPath, saveOptions);

            // ----- Verification -----
            // Load the generated HTML as text
            string htmlContent = File.ReadAllText(htmlPath);

            // Check that the hidden sheet name does NOT appear in the HTML
            bool hiddenSheetPresent = htmlContent.Contains("HiddenSheet");

            Console.WriteLine($"Hidden sheet present in HTML: {hiddenSheetPresent}");
            Console.WriteLine(hiddenSheetPresent
                ? "Verification FAILED: Hidden worksheet was exported."
                : "Verification PASSED: Hidden worksheet is absent from the HTML.");

            // For completeness, demonstrate the opposite case (export hidden worksheets)
            saveOptions.ExportHiddenWorksheet = true;
            string htmlPathWithHidden = "output_with_hidden.html";
            workbook.Save(htmlPathWithHidden, saveOptions);

            string htmlContentWithHidden = File.ReadAllText(htmlPathWithHidden);
            bool hiddenSheetNowPresent = htmlContentWithHidden.Contains("HiddenSheet");

            Console.WriteLine($"Hidden sheet present after enabling export: {hiddenSheetNowPresent}");
        }
    }
}
