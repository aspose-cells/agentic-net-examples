// Title: Export Workbook to HTML without Hidden Rows and Optional Hidden Worksheet Inclusion – Aspose.Cells for .NET
// Description: This C# example demonstrates how to create a workbook, hide a specific row and an entire worksheet, and then save the file to HTML using Aspose.Cells. By setting HtmlSaveOptions.HiddenRowDisplayType to Remove, hidden rows are omitted from the output. The ExportHiddenWorksheet flag is toggled to produce two HTML files: one that excludes hidden worksheets and another that includes them. The code also shows how to verify the results by checking the generated HTML content.
// Keywords: Aspose.Cells HTML export | Hide row in HTML output | HtmlSaveOptions HiddenRowDisplayType Remove | ExportHiddenWorksheet property | C# export Excel to HTML | remove hidden rows Aspose.Cells | .NET workbook to HTML | exclude hidden worksheets HTML | Aspose.Cells hidden sheet export
// Common Searches: Aspose.Cells hide row when exporting to HTML | How to exclude hidden rows in HTML output using Aspose.Cells | ExportHiddenWorksheet option Aspose.Cells .NET | Remove hidden rows from HTML with HtmlSaveOptions | Include hidden worksheets in HTML export Aspose.Cells
// Developer Intent: Generate HTML from an Excel workbook that excludes hidden rows and optionally includes hidden worksheets based on the ExportHiddenWorksheet setting.
// Use Cases: Create a clean HTML report that shows only visible data from an Excel file. | Produce two versions of an HTML export—one for public sharing without hidden worksheets and another for internal review that retains hidden sheets. | Programmatically validate the HTML output by confirming hidden rows are absent and hidden worksheet content appears only when desired.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to HTML, removing hidden rows and toggling ExportHiddenWorksheet for different outputs. | Generate a unit test that loads the saved HTML files and asserts that hidden rows are not present while hidden worksheet data appears only when ExportHiddenWorksheet is true. | Explain how HtmlHiddenRowDisplayType.Remove works during HTML conversion and how it interacts with the ExportHiddenWorksheet flag.

using System;
using System.IO;
using Aspose.Cells;

// This C# example demonstrates how to create a workbook, hide a specific row and an entire worksheet, and then save the file to HTML using Aspose.Cells. By setting HtmlSaveOptions.HiddenRowDisplayType to Remove, hidden rows are omitted from the output. The ExportHiddenWorksheet flag is toggled to produce two HTML files: one that excludes hidden worksheets and another that includes them. The code also shows how to verify the results by checking the generated HTML content.
class HiddenRowsHtmlDemo
{
    static void Main()
    {
        // Create a new workbook and get the first (visible) worksheet
        Workbook workbook = new Workbook();
        Worksheet visibleSheet = workbook.Worksheets[0];
        visibleSheet.Name = "VisibleSheet";

        // Add data: two visible rows and one hidden row
        visibleSheet.Cells["A1"].PutValue("Visible Row 1");
        visibleSheet.Cells["A2"].PutValue("Hidden Row");
        visibleSheet.Cells["A3"].PutValue("Visible Row 2");

        // Hide the second row (index 1)
        visibleSheet.Cells.HideRow(1);

        // Add a second worksheet and hide it
        Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
        hiddenSheet.Cells["A1"].PutValue("Data in hidden sheet");
        hiddenSheet.IsVisible = false; // Mark worksheet as hidden

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Remove hidden rows from the HTML output
            HiddenRowDisplayType = HtmlHiddenRowDisplayType.Remove
        };

        // -------------------------------------------------
        // Save HTML without exporting hidden worksheets
        // -------------------------------------------------
        htmlOptions.ExportHiddenWorksheet = false;
        string pathWithoutHiddenSheet = "output_without_hidden_sheet.html";
        workbook.Save(pathWithoutHiddenSheet, htmlOptions);

        // -------------------------------------------------
        // Save HTML with hidden worksheets exported
        // -------------------------------------------------
        htmlOptions.ExportHiddenWorksheet = true;
        string pathWithHiddenSheet = "output_with_hidden_sheet.html";
        workbook.Save(pathWithHiddenSheet, htmlOptions);

        // Load the generated HTML files for verification
        string htmlWithout = File.ReadAllText(pathWithoutHiddenSheet);
        string htmlWith = File.ReadAllText(pathWithHiddenSheet);

        // Verify that the hidden row is omitted in both files
        Console.WriteLine("Hidden row present in file without hidden sheet? " + (htmlWithout.Contains("Hidden Row") ? "Yes" : "No"));
        Console.WriteLine("Hidden row present in file with hidden sheet? " + (htmlWith.Contains("Hidden Row") ? "Yes" : "No"));

        // Verify that the hidden worksheet content follows ExportHiddenWorksheet setting
        Console.WriteLine("Hidden worksheet data present in file without hidden sheet? " + (htmlWithout.Contains("Data in hidden sheet") ? "Yes" : "No"));
        Console.WriteLine("Hidden worksheet data present in file with hidden sheet? " + (htmlWith.Contains("Data in hidden sheet") ? "Yes" : "No"));
    }
}
