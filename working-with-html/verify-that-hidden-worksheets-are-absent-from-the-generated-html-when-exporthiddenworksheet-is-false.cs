// Title: Exclude Hidden Worksheets from HTML Export with Aspose.Cells (C#)
// Description: Creates a workbook with a visible and a hidden sheet, sets HtmlSaveOptions.ExportHiddenWorksheet to false, saves to HTML, reads the file, and confirms that only the visible sheet appears in the output.
// Keywords: Aspose.Cells hidden worksheet HTML export | ExportHiddenWorksheet false C# | verify hidden sheet not in HTML | Aspose.Cells HtmlSaveOptions example | C# hide worksheet during HTML conversion
// Common Searches: Aspose.Cells prevent hidden sheets from exporting to HTML | C# check hidden worksheet excluded from HTML output | HtmlSaveOptions ExportHiddenWorksheet usage | how to hide worksheets in Aspose.Cells HTML export | verify hidden worksheet omission in generated HTML
// Developer Intent: Confirm that hidden worksheets are omitted from the HTML file when ExportHiddenWorksheet is set to false.
// Use Cases: Publish web‑ready reports that contain only visible data. | Automated tests to ensure confidential sheets are not exposed in HTML. | Generate documentation where hidden worksheets must remain private.
// AI Prompts: Write a C# unit test using Aspose.Cells that asserts hidden worksheets are not present in the HTML output when ExportHiddenWorksheet is false. | Show how to log a warning if a hidden worksheet is exported despite ExportHiddenWorksheet being false. | Explain how to enumerate visible worksheets before saving a workbook to HTML with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHiddenWorksheetVerification
{
    // Creates a workbook with a visible and a hidden sheet, sets HtmlSaveOptions.ExportHiddenWorksheet to false, saves to HTML, reads the file, and confirms that only the visible sheet appears in the output.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Configure the first (visible) worksheet
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Data in visible sheet");

            // Add a second worksheet and hide it
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Data in hidden sheet");
            hiddenSheet.IsVisible = false; // Mark as hidden

            // Set HTML save options to exclude hidden worksheets
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = false, // Do not export hidden worksheets
                ExportActiveWorksheetOnly = false // Export the whole workbook (visible sheets only)
            };

            // Define output HTML file path
            string htmlPath = "WorkbookWithoutHiddenSheet.html";

            // Save the workbook to HTML using the specified options
            workbook.Save(htmlPath, saveOptions);

            // Load the generated HTML content for verification
            string htmlContent = File.ReadAllText(htmlPath);

            // Verify that the hidden worksheet name does NOT appear in the HTML
            bool hiddenSheetPresent = htmlContent.Contains(hiddenSheet.Name);
            Console.WriteLine($"Hidden worksheet \"{hiddenSheet.Name}\" present in HTML: {hiddenSheetPresent}");

            // Verify that the visible worksheet name DOES appear in the HTML
            bool visibleSheetPresent = htmlContent.Contains(visibleSheet.Name);
            Console.WriteLine($"Visible worksheet \"{visibleSheet.Name}\" present in HTML: {visibleSheetPresent}");

            // Output verification result
            if (!hiddenSheetPresent && visibleSheetPresent)
            {
                Console.WriteLine("Verification passed: hidden worksheets are absent from the generated HTML.");
            }
            else
            {
                Console.WriteLine("Verification failed: hidden worksheets were found in the generated HTML.");
            }
        }
    }
}
