// Title: Error handling for missing hidden worksheets with ExportHiddenWorksheet=true in AspNet Aspose.Cells HTML export (C#)
// Description: Demonstrates how to validate hidden worksheets before saving a workbook as HTML when ExportHiddenWorksheet is enabled. The sample checks for hidden sheets without a name, throws an InvalidOperationException, and logs the error to prevent runtime failures.
// Keywords: Aspose.Cells | ExportHiddenWorksheet | HTML export | hidden worksheet validation | C# error handling | Workbook.Save exception | missing sheet name | Aspose.Cells HTMLSaveOptions | Excel to HTML conversion
// Common Searches: Aspose.Cells ExportHiddenWorksheet missing sheet name | C# validate hidden worksheets before HTML export | How to catch errors when exporting hidden sheets with Aspose.Cells | HTMLSaveOptions ExportHiddenWorksheet exception handling | Aspose.Cells hidden worksheet validation example
// Developer Intent: Ensure hidden worksheets have valid names before exporting to HTML when ExportHiddenWorksheet is true, and provide graceful error handling.
// Use Cases: Prevent crashes by detecting hidden sheets without a name prior to HTML conversion. | Log detailed information about corrupt hidden worksheets for troubleshooting. | Automatically rename, skip, or report hidden worksheets that lack a name to guarantee successful HTML output.
// AI Prompts: Generate a C# method that scans a Workbook for hidden worksheets with empty names and throws an InvalidOperationException before HTML export using Aspose.Cells. | Create C# code that logs the index of any hidden worksheet missing a name, then continues exporting the remaining sheets to HTML with ExportHiddenWorksheet=true. | Write a unit test in C# that verifies ValidateWorksheetsForExport throws an exception when a hidden worksheet has no name in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

// Demonstrates how to validate hidden worksheets before saving a workbook as HTML when ExportHiddenWorksheet is enabled. The sample checks for hidden sheets without a name, throws an InvalidOperationException, and logs the error to prevent runtime failures.
class ExportHiddenWorksheetDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();

        // First (visible) worksheet
        Worksheet visibleSheet = workbook.Worksheets[0];
        visibleSheet.Name = "VisibleSheet";
        visibleSheet.Cells["A1"].PutValue("Visible Data");

        // Add a hidden worksheet
        Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
        hiddenSheet.Cells["A1"].PutValue("Hidden Data");
        hiddenSheet.IsVisible = false; // hide the sheet

        // Configure HTML save options to export hidden worksheets
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportHiddenWorksheet = true,
            ExportActiveWorksheetOnly = false
        };

        // Perform validation and handle any errors before saving
        try
        {
            ValidateWorksheetsForExport(workbook, saveOptions);
            workbook.Save("output.html", saveOptions);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            // Log the error – in a real application you might log to a file or monitoring system
            Console.WriteLine($"Error during export: {ex.Message}");
        }
    }

    // Checks that hidden worksheets are valid when ExportHiddenWorksheet is enabled
    static void ValidateWorksheetsForExport(Workbook wb, HtmlSaveOptions options)
    {
        if (!options.ExportHiddenWorksheet)
            return; // No validation needed when the option is disabled

        foreach (Worksheet ws in wb.Worksheets)
        {
            // If a worksheet is marked as hidden but lacks a name, treat it as missing/corrupt
            if (!ws.IsVisible && string.IsNullOrEmpty(ws.Name))
            {
                throw new InvalidOperationException(
                    $"Hidden worksheet at index {ws.Index} is missing a valid name.");
            }
        }
    }
}
