// Title: Export a workbook to HTML without hidden worksheets using Aspose.Cells in C#
// AI Prompts: Write C# that creates a workbook, hides a specific worksheet, and saves it to HTML while disabling hidden worksheet export. | Provide C# code that loads the generated HTML file and verifies that the hidden worksheet name does not appear. | Demonstrate configuring Aspose.Cells HtmlSaveOptions to exclude hidden sheets during HTML conversion.
// Common Searches: Aspose.Cells C# export workbook to HTML while omitting hidden sheets | How to set ExportHiddenWorksheet to false in Aspose.Cells HTML conversion | HtmlSaveOptions example for excluding hidden worksheets in .NET | Check that hidden worksheet is not included in generated HTML with Aspose.Cells | C# code to hide a worksheet and save workbook as HTML using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet | C# hide worksheet before HTML export | HTML export without hidden sheets | validate hidden worksheet omission Aspose.Cells | Aspose.Cells workbook to HTML excluding hidden sheets

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook with one visible and one hidden worksheet, configures HtmlSaveOptions.ExportHiddenWorksheet = false, saves the workbook as HTML, reads the output file, and confirms that the hidden sheet name is absent while the visible sheet name is present.
class ExportHiddenWorksheetDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add first visible worksheet and put some data
        Worksheet visibleSheet = workbook.Worksheets[0];
        visibleSheet.Name = "VisibleSheet";
        visibleSheet.Cells["A1"].PutValue("Visible Data");

        // Add a second worksheet and hide it
        int hiddenIndex = workbook.Worksheets.Add();
        Worksheet hiddenSheet = workbook.Worksheets[hiddenIndex];
        hiddenSheet.Name = "HiddenSheet";
        hiddenSheet.Cells["A1"].PutValue("Hidden Data");
        hiddenSheet.IsVisible = false; // Hide the worksheet

        // Set ExportHiddenWorksheet to false so hidden sheets are omitted from HTML
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportHiddenWorksheet = false;

        // Define output HTML file path
        string htmlPath = "ExportedWorkbook.html";

        // Save the workbook as HTML
        workbook.Save(htmlPath, saveOptions);

        // Verify that the hidden sheet is omitted from the resulting HTML
        string htmlContent = File.ReadAllText(htmlPath);

        bool hiddenSheetPresent = htmlContent.Contains("HiddenSheet");
        bool visibleSheetPresent = htmlContent.Contains("VisibleSheet");

        Console.WriteLine("Hidden sheet present in HTML: " + hiddenSheetPresent);
        Console.WriteLine("Visible sheet present in HTML: " + visibleSheetPresent);

        // Expected output:
        // Hidden sheet present in HTML: False
        // Visible sheet present in HTML: True
    }
}
