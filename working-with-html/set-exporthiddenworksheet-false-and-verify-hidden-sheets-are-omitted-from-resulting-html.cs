// Title: Export an Aspose.Cells workbook to HTML while excluding hidden worksheets in C#
// AI Prompts: Write C# code that uses Aspose.Cells to save a workbook as HTML and sets HtmlSaveOptions.ExportHiddenWorksheet to false so hidden sheets are not included. | Add C# logic that reads the generated HTML file and checks that the name of any hidden worksheet does not appear, outputting a verification result.
// Common Searches: Aspose.Cells C# export workbook to HTML without hidden sheets | How to prevent hidden worksheets from being saved in HTML using Aspose.Cells | C# code example for HtmlSaveOptions ExportHiddenWorksheet false | Verify hidden worksheet exclusion in Aspose.Cells HTML output programmatically | Skip hidden worksheets when converting Excel to HTML with Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet C# | C# export workbook to HTML without hidden worksheets | verify hidden sheet omission in generated HTML | Aspose.Cells HTML conversion exclude hidden sheets | programmatic check hidden worksheet presence in HTML

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook with a visible and a hidden sheet, configures HtmlSaveOptions.ExportHiddenWorksheet = false, saves the workbook as HTML, reads the output file, and confirms that the hidden sheet name is absent, demonstrating how to omit hidden worksheets from HTML export using Aspose.Cells for .NET.
class ExportHiddenWorksheetDemo
{
    static void Main()
    {
        // Create a new workbook with two worksheets
        Workbook workbook = new Workbook();
        Worksheet visibleSheet = workbook.Worksheets[0];
        visibleSheet.Name = "VisibleSheet";

        // Add a second worksheet and hide it
        int hiddenIndex = workbook.Worksheets.Add();
        Worksheet hiddenSheet = workbook.Worksheets[hiddenIndex];
        hiddenSheet.Name = "HiddenSheet";
        hiddenSheet.IsVisible = false; // Mark the sheet as hidden

        // Configure HTML save options to exclude hidden worksheets
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHiddenWorksheet = false; // Do not export hidden sheets

        // Save the workbook to HTML
        string htmlPath = "output.html";
        workbook.Save(htmlPath, htmlOptions);

        // Verify that the hidden sheet name does not appear in the generated HTML
        string htmlContent = File.ReadAllText(htmlPath);
        bool hiddenSheetPresent = htmlContent.Contains("HiddenSheet", StringComparison.OrdinalIgnoreCase);
        Console.WriteLine(hiddenSheetPresent
            ? "Verification failed: hidden sheet was exported."
            : "Verification succeeded: hidden sheet was omitted.");
    }
}
