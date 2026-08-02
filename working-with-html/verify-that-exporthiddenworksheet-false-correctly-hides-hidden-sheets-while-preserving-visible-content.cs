// Title: Aspose.Cells .NET: ExportHiddenWorksheet false hides hidden sheets in HTML export
// Description: Shows how HtmlSaveOptions.ExportHiddenWorksheet can be set to false to omit hidden worksheets when saving a workbook to HTML, then switched to true to include them, while preserving all visible sheet content.
// Keywords: Aspose.Cells | ExportHiddenWorksheet | HTML export | .NET | C# | hide hidden worksheets | visible worksheet | HtmlSaveOptions | Workbook to HTML | ExportActiveWorksheetOnly | Aspose.Cells example
// Common Searches: Aspose.Cells hide hidden worksheets in HTML export | ExportHiddenWorksheet false example C# | HTML save options Aspose.Cells exclude hidden sheets | How to prevent hidden sheets from appearing in HTML output | Toggle ExportHiddenWorksheet to include hidden worksheets
// Developer Intent: Confirm that ExportHiddenWorksheet = false excludes hidden worksheets from the HTML output while still exporting all visible sheets.
// Use Cases: Generate an HTML report that shows only the worksheets intended for public view. | Validate that hidden data remains confidential by excluding it from exported HTML files. | Switch between excluding and including hidden sheets without changing other save options.
// AI Prompts: Create a unit test in C# using Aspose.Cells that asserts hidden worksheets are not present in the HTML file when ExportHiddenWorksheet is false. | Write a script that lists which worksheets will be exported based on the current HtmlSaveOptions settings. | Explain the interaction between ExportHiddenWorksheet and ExportActiveWorksheetOnly when exporting a workbook to HTML with Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how HtmlSaveOptions.ExportHiddenWorksheet can be set to false to omit hidden worksheets when saving a workbook to HTML, then switched to true to include them, while preserving all visible sheet content.
class ExportHiddenWorksheetDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Configure the first (default) worksheet as visible and add data
        Worksheet visibleSheet = workbook.Worksheets[0];
        visibleSheet.Name = "VisibleSheet";
        visibleSheet.Cells["A1"].PutValue("Visible Data");

        // Add a second worksheet and hide it
        Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
        hiddenSheet.Cells["A1"].PutValue("Hidden Data");
        hiddenSheet.IsVisible = false; // Mark the sheet as hidden

        // Prepare HTML save options: do NOT export hidden worksheets
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportHiddenWorksheet = false,   // Hide hidden sheets in the output
            ExportActiveWorksheetOnly = false // Export the whole workbook
        };

        // Save the workbook to HTML; only the visible sheet will be present
        workbook.Save("output_without_hidden.html", saveOptions);

        // Change the option to include hidden worksheets
        saveOptions.ExportHiddenWorksheet = true;

        // Save again; both visible and hidden sheets will be exported
        workbook.Save("output_with_hidden.html", saveOptions);
    }
}
