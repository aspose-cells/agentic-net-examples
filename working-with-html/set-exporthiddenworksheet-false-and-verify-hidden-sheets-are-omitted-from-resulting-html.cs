// Title: Export HTML without Hidden Worksheets using Aspose.Cells for .NET
// Description: Shows how to configure HtmlSaveOptions.ExportHiddenWorksheet = false so hidden sheets are excluded from the HTML output, verifies the result, and demonstrates re‑enabling the option for comparison.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportHiddenWorksheet | hide worksheets | HTML export | C# | .NET | workbook to HTML | exclude hidden sheets | verify HTML output
// Common Searches: Aspose.Cells export hidden sheet to HTML | C# HtmlSaveOptions exclude hidden worksheets | how to omit hidden worksheets when saving as HTML | verify hidden sheet not in HTML output Aspose | ExportHiddenWorksheet false example
// Developer Intent: Create an HTML file that contains only the visible worksheets of a workbook.
// Use Cases: Publish a public‑facing HTML report that must not reveal data from hidden tabs. | Automated testing that confirms hidden worksheet content is absent from the generated HTML. | Switch between including and excluding hidden sheets to debug layout differences.
// AI Prompts: Generate C# code with Aspose.Cells that exports a workbook to HTML while skipping hidden worksheets and validates the output. | Provide a reusable method that accepts a Workbook and a bool for ExportHiddenWorksheet, saves HTML, and returns a verification flag. | Explain how ExportHiddenWorksheet interacts with ExportActiveWorksheetOnly during HTML conversion in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHiddenSheetExportDemo
{
    // Shows how to configure HtmlSaveOptions.ExportHiddenWorksheet = false so hidden sheets are excluded from the HTML output, verifies the result, and demonstrates re‑enabling the option for comparison.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ----- Visible worksheet -----
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Visible Data");

            // ----- Hidden worksheet -----
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Hidden Data");
            hiddenSheet.IsVisible = false; // Mark the sheet as hidden

            // Configure HTML save options to exclude hidden worksheets
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = false, // Do not export hidden worksheets
                ExportActiveWorksheetOnly = false // Export the whole workbook (visible sheets only)
            };

            // Save the workbook to HTML
            string htmlPath = "output_without_hidden.html";
            workbook.Save(htmlPath, saveOptions);

            // Verify that the hidden sheet's content is not present in the generated HTML
            string htmlContent = File.ReadAllText(htmlPath);
            bool hiddenDataFound = htmlContent.Contains("Hidden Data");
            bool visibleDataFound = htmlContent.Contains("Visible Data");

            Console.WriteLine($"Visible data present: {visibleDataFound}");
            Console.WriteLine($"Hidden data present: {hiddenDataFound}");

            if (!hiddenDataFound && visibleDataFound)
            {
                Console.WriteLine("Verification succeeded: hidden worksheet was omitted from the HTML output.");
            }
            else
            {
                Console.WriteLine("Verification failed: hidden worksheet was not correctly omitted.");
            }

            // Optional: demonstrate the opposite setting (export hidden worksheets)
            saveOptions.ExportHiddenWorksheet = true;
            string htmlPathWithHidden = "output_with_hidden.html";
            workbook.Save(htmlPathWithHidden, saveOptions);

            // Verify that hidden data now appears
            string htmlContentWithHidden = File.ReadAllText(htmlPathWithHidden);
            bool hiddenDataNowFound = htmlContentWithHidden.Contains("Hidden Data");
            Console.WriteLine($"After enabling ExportHiddenWorksheet, hidden data present: {hiddenDataNowFound}");
        }
    }
}
