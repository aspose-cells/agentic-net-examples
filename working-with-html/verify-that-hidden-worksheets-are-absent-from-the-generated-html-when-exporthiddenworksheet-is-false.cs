// Title: Check that hidden worksheets are omitted from HTML output when ExportHiddenWorksheet is set to false using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that creates a visible sheet and a hidden sheet, saves the workbook to HTML with HtmlSaveOptions.ExportHiddenWorksheet = false, then reads the HTML file and confirms the hidden sheet name is absent. | Write a C# unit test that asserts Aspose.Cells does not export hidden worksheets to HTML when the ExportHiddenWorksheet option is disabled.
// Common Searches: Aspose.Cells how to prevent hidden sheets from appearing in exported HTML | C# HtmlSaveOptions ExportHiddenWorksheet false example | verify hidden worksheet exclusion from HTML using Aspose.Cells .NET | unit test Aspose.Cells HTML export without hidden worksheets | exclude hidden worksheets when saving workbook as HTML Aspose.Cells
// Tags: Aspose.Cells hide worksheet HTML export | HtmlSaveOptions ExportHiddenWorksheet false | C# verify hidden sheet omission from HTML | Aspose.Cells generate HTML without hidden sheets | unit testing hidden worksheet exclusion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook with one visible and one hidden worksheet, configures HtmlSaveOptions.ExportHiddenWorksheet to false, saves the workbook as HTML, reads the resulting file, and validates that the visible sheet name is present while the hidden sheet name is not.
class Program
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Configure the first worksheet as visible
        Worksheet visibleSheet = workbook.Worksheets[0];
        visibleSheet.Name = "VisibleSheet";
        visibleSheet.Cells["A1"].PutValue("Visible Data");

        // Add a second worksheet and hide it
        Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
        hiddenSheet.Cells["A1"].PutValue("Hidden Data");
        hiddenSheet.IsVisible = false; // hide the worksheet

        // Export to HTML with ExportHiddenWorksheet set to false (save rule)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHiddenWorksheet = false;
        string htmlPath = "output.html";
        workbook.Save(htmlPath, htmlOptions); // save rule

        // Load the generated HTML file (load rule) and verify content
        string htmlContent = File.ReadAllText(htmlPath);
        bool visiblePresent = htmlContent.Contains("VisibleSheet");
        bool hiddenPresent = htmlContent.Contains("HiddenSheet");

        Console.WriteLine($"Visible sheet present: {visiblePresent}");
        Console.WriteLine($"Hidden sheet present: {hiddenPresent}");

        if (visiblePresent && !hiddenPresent)
        {
            Console.WriteLine("Test passed: hidden worksheet is absent from the HTML.");
        }
        else
        {
            Console.WriteLine("Test failed: hidden worksheet was found in the HTML.");
        }
    }
}
