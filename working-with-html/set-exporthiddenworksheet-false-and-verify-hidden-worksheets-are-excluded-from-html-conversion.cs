// Title: C# – Export Workbook to HTML Without Hidden Worksheets Using Aspose.Cells
// Description: Shows how to build a workbook with one visible and one hidden sheet, set HtmlSaveOptions.ExportHiddenWorksheet = false to exclude the hidden sheet from the HTML output, read the file to confirm the omission, and then repeat the export with ExportHiddenWorksheet = true for comparison.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportHiddenWorksheet | C# .NET | hide worksheet HTML export | exclude hidden sheets | Workbook to HTML | visible worksheets only | Aspose.Cells example
// Common Searches: Aspose.Cells hide hidden worksheets in HTML | ExportHiddenWorksheet false C# | How to skip hidden sheets when saving as HTML Aspose.Cells | Verify hidden sheet not in HTML output Aspose | Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly vs ExportHiddenWorksheet
// Developer Intent: Export a workbook to HTML while omitting any worksheets that are marked as hidden.
// Use Cases: Create an HTML report that displays only the data intended for end‑users. | Generate separate public and internal HTML files by toggling the hidden‑sheet export option. | Automate a validation step that ensures hidden worksheets are not rendered in the final HTML.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to HTML excluding hidden worksheets and confirm the result. | Provide a C# unit test that asserts ExportHiddenWorksheet = false removes hidden sheet content from the saved HTML file. | Explain the interaction between ExportActiveWorksheetOnly and ExportHiddenWorksheet when converting a workbook to HTML.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExportHiddenWorksheetDemo
{
    // Shows how to build a workbook with one visible and one hidden sheet, set HtmlSaveOptions.ExportHiddenWorksheet = false to exclude the hidden sheet from the HTML output, read the file to confirm the omission, and then repeat the export with ExportHiddenWorksheet = true for comparison.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the default worksheet (visible)
            Workbook workbook = new Workbook();
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Data in visible sheet");

            // Add a second worksheet and hide it
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Data in hidden sheet");
            hiddenSheet.IsVisible = false; // Mark worksheet as hidden

            // Configure HTML save options to exclude hidden worksheets
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = false, // Do not export hidden worksheets
                ExportActiveWorksheetOnly = false // Export the whole workbook (visible sheets only)
            };

            // Define output path
            string outputPath = Path.Combine(Environment.CurrentDirectory, "WorkbookWithoutHidden.html");

            // Save the workbook to HTML; hidden sheet will be omitted
            workbook.Save(outputPath, saveOptions);

            // Optional: Verify that the hidden worksheet content is not present in the HTML file
            // (In a real test, you would read the file and assert that "HiddenSheet" or its data is absent)
            string htmlContent = File.ReadAllText(outputPath);
            bool hiddenSheetPresent = htmlContent.Contains("HiddenSheet") || htmlContent.Contains("Data in hidden sheet");
            Console.WriteLine($"Hidden worksheet exported? {(hiddenSheetPresent ? "Yes" : "No")}");

            // For comparison, export with hidden worksheets included
            saveOptions.ExportHiddenWorksheet = true;
            string outputPathWithHidden = Path.Combine(Environment.CurrentDirectory, "WorkbookWithHidden.html");
            workbook.Save(outputPathWithHidden, saveOptions);
            Console.WriteLine("Export completed. Check the generated HTML files.");
        }
    }
}
