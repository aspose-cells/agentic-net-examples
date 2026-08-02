// Title: Exclude hidden worksheets from HTML export with Aspose.Cells for .NET (ExportHiddenWorksheet = false)
// Description: Demonstrates how to create a workbook containing a visible and a hidden sheet, configure HtmlSaveOptions.ExportHiddenWorksheet to false, save the file as HTML, and programmatically confirm that the hidden sheet's content is omitted from the generated markup.
// Keywords: Aspose.Cells | .NET | C# | HtmlSaveOptions | ExportHiddenWorksheet | hide worksheet | HTML export | exclude hidden sheet | visible worksheets only | workbook to HTML example
// Common Searches: Aspose.Cells hide worksheet in HTML export | ExportHiddenWorksheet false C# example | skip hidden sheets when saving as HTML Aspose.Cells | verify hidden sheet not present in HTML output | HtmlSaveOptions ExportHiddenWorksheet usage
// Developer Intent: The developer needs to prevent hidden worksheets from appearing in the HTML representation of a workbook.
// Use Cases: Generate web‑ready reports that contain only data from visible sheets. | Protect confidential information by omitting hidden tabs during HTML conversion. | Automate validation of HTML output to ensure hidden content is not leaked.
// AI Prompts: Show how to modify the sample to export only the active worksheet while ignoring hidden ones. | Provide a C# unit test that asserts the hidden sheet's text is absent from the saved HTML file. | Explain the impact on the HTML result when ExportHiddenWorksheet is set to true versus false.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a workbook containing a visible and a hidden sheet, configure HtmlSaveOptions.ExportHiddenWorksheet to false, save the file as HTML, and programmatically confirm that the hidden sheet's content is omitted from the generated markup.
class ExportHiddenWorksheetDemo
{
    static void Main()
    {
        // Create a new workbook and add data to a visible sheet
        Workbook workbook = new Workbook();
        Worksheet visibleSheet = workbook.Worksheets[0];
        visibleSheet.Name = "VisibleSheet";
        visibleSheet.Cells["A1"].PutValue("Visible Data");

        // Add a hidden worksheet with its own data
        Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
        hiddenSheet.Cells["A1"].PutValue("Hidden Data");
        hiddenSheet.IsVisible = false; // Mark the sheet as hidden

        // Configure HTML save options to exclude hidden worksheets
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportHiddenWorksheet = false, // Do not export hidden worksheets
            ExportActiveWorksheetOnly = false
        };

        // Save the workbook to HTML; hidden sheet should not appear
        string htmlPath = "output_without_hidden.html";
        workbook.Save(htmlPath, saveOptions);

        // Verify that the hidden worksheet's content is not present in the HTML file
        string htmlContent = File.ReadAllText(htmlPath);
        bool hiddenDataFound = htmlContent.Contains("Hidden Data");
        Console.WriteLine("Hidden worksheet excluded from HTML: " + (!hiddenDataFound));
    }
}
