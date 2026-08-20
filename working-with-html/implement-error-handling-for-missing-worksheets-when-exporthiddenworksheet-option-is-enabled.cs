// Title: C# – Validate hidden worksheets before HTML export with ExportHiddenWorksheet in Aspose.Cells
// Description: This Aspose.Cells for .NET example creates a workbook containing a visible and a hidden sheet, enables HtmlSaveOptions.ExportHiddenWorksheet, and runs a validation routine that checks for hidden worksheets. If no hidden sheet is found, an InvalidOperationException is thrown and caught, demonstrating robust error handling for the missing‑worksheet scenario.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | ExportHiddenWorksheet | hidden worksheet validation | HTML export error handling | InvalidOperationException | missing hidden sheet | workbook export
// Common Searches: Aspose.Cells ExportHiddenWorksheet validation example | how to check for hidden worksheets before saving HTML | C# error handling when hidden sheet is missing in Aspose.Cells | ExportHiddenWorksheet throws exception if no hidden sheets | validate hidden worksheets Aspose.Cells HtmlSaveOptions
// Developer Intent: Add a safeguard that raises an exception when ExportHiddenWorksheet is true but the workbook lacks hidden worksheets.
// Use Cases: Prevent silent HTML output when hidden sheets are required for compliance or reporting. | Show a clear error message in UI or logs when the hidden‑sheet export option cannot be satisfied. | Filter or flag workbooks in batch jobs that miss hidden worksheets while ExportHiddenWorksheet is enabled.
// AI Prompts: Generate a C# method that scans a Workbook for hidden worksheets and throws InvalidOperationException when none are present and HtmlSaveOptions.ExportHiddenWorksheet is true. | Write NUnit tests for ValidateWorksheetsForExport covering: hidden sheet exists, hidden sheet absent, and ExportHiddenWorksheet disabled. | Refactor the validation to log the missing‑worksheet error using Aspose.Cells logging and continue processing a collection of workbooks.

using System;
using Aspose.Cells;

// This Aspose.Cells for .NET example creates a workbook containing a visible and a hidden sheet, enables HtmlSaveOptions.ExportHiddenWorksheet, and runs a validation routine that checks for hidden worksheets. If no hidden sheet is found, an InvalidOperationException is thrown and caught, demonstrating robust error handling for the missing‑worksheet scenario.
class ExportHiddenWorksheetDemo
{
    static void Main()
    {
        // -------------------- Create a workbook with visible and hidden sheets --------------------
        Workbook workbook = new Workbook();

        // First (visible) worksheet
        Worksheet visibleSheet = workbook.Worksheets[0];
        visibleSheet.Name = "VisibleSheet";
        visibleSheet.Cells["A1"].PutValue("Visible Data");

        // Add a hidden worksheet
        Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
        hiddenSheet.Cells["A1"].PutValue("Hidden Data");
        hiddenSheet.IsVisible = false; // hide the sheet

        // -------------------- Configure HTML save options --------------------
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportHiddenWorksheet = true,   // request exporting hidden worksheets
            ExportActiveWorksheetOnly = false
        };

        // -------------------- Save with validation and error handling --------------------
        try
        {
            ValidateWorksheetsForExport(workbook, saveOptions);
            workbook.Save("output_with_hidden.html", saveOptions);
            Console.WriteLine("Workbook saved successfully with hidden worksheets exported.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during export: {ex.Message}");
        }

        // -------------------- Simulate a missing hidden worksheet scenario --------------------
        // Remove the hidden sheet to create a "missing worksheet" condition
        try
        {
            // RemoveAt uses the sheet's index; ensure the sheet still exists before removal
            if (hiddenSheet != null && hiddenSheet.Index >= 0 && hiddenSheet.Index < workbook.Worksheets.Count)
            {
                workbook.Worksheets.RemoveAt(hiddenSheet.Index);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error removing hidden sheet: {ex.Message}");
        }

        // Attempt export again; expect validation to fail because no hidden sheets exist
        try
        {
            ValidateWorksheetsForExport(workbook, saveOptions);
            workbook.Save("output_missing_hidden.html", saveOptions);
        }
        catch (Exception ex)
        {
            // Expected: ExportHiddenWorksheet is true but no hidden worksheets exist
            Console.WriteLine($"Handled missing worksheet case: {ex.Message}");
        }
    }

    // Validates that hidden worksheets exist when ExportHiddenWorksheet option is enabled
    static void ValidateWorksheetsForExport(Workbook workbook, HtmlSaveOptions options)
    {
        if (options.ExportHiddenWorksheet)
        {
            bool hasHidden = false;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (!ws.IsVisible)
                {
                    hasHidden = true;
                    break;
                }
            }

            if (!hasHidden)
            {
                throw new InvalidOperationException(
                    "ExportHiddenWorksheet is enabled, but the workbook does not contain any hidden worksheets.");
            }
        }
    }
}
