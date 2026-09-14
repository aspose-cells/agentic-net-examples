// Title: Include hidden worksheets when exporting a workbook to HTML with Aspose.Cells for .NET and verify the output
// AI Prompts: Generate C# code that creates a workbook, hides a worksheet, enables HtmlSaveOptions.ExportHiddenWorksheet, and saves the workbook as an HTML file. | Add C# logic to read the generated HTML file and confirm that the hidden worksheet's data appears in the HTML content.
// Common Searches: Aspose.Cells .NET export hidden worksheet to HTML example | How to include hidden Excel sheets when saving as HTML using HtmlSaveOptions | C# verify hidden sheet content in generated HTML with Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet property usage guide | Check hidden worksheet inclusion after HTML conversion Aspose.Cells
// Tags: Aspose.Cells export hidden worksheets to HTML | HtmlSaveOptions ExportHiddenWorksheet property | C# verify hidden sheet content in HTML output | save workbook as HTML including hidden sheets | detect hidden worksheet data in generated HTML

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook with a visible and a hidden sheet, sets HtmlSaveOptions.ExportHiddenWorksheet to true, saves the workbook as HTML, reads the resulting file, checks for the hidden sheet's content, and prints whether the hidden sheet was successfully included.
class ExportHiddenWorksheetExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the default (first) worksheet and add some visible content
        Worksheet visibleSheet = workbook.Worksheets[0];
        visibleSheet.Name = "VisibleSheet";
        visibleSheet.Cells["A1"].PutValue("Visible Content");

        // Add a second worksheet and hide it
        Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
        hiddenSheet.Cells["A1"].PutValue("Hidden Content");
        hiddenSheet.IsVisible = false; // Mark the sheet as hidden

        // Configure HTML save options to export hidden worksheets
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHiddenWorksheet = true;

        // Define output HTML file path
        string htmlPath = "output.html";

        // Save the workbook as HTML with the specified options
        workbook.Save(htmlPath, htmlOptions);

        // Verify that the hidden sheet's content is present in the generated HTML
        string htmlContent = File.ReadAllText(htmlPath);
        bool hiddenIncluded = htmlContent.Contains("Hidden Content");

        Console.WriteLine("Hidden sheet included in HTML: " + hiddenIncluded);
    }
}
