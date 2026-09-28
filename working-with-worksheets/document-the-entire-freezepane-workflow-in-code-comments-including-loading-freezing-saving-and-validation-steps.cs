// Title: Implement a complete freeze‑pane workflow with Aspose.Cells for .NET: load workbook, apply FreezePanes, verify, unfreeze, and save
// AI Prompts: Generate C# code that loads an Excel file using Aspose.Cells, freezes panes at cell E5 with 4 rows and 2 columns, retrieves the frozen row/column indices via GetFreezedPanes, and saves the result. | Provide a C# example that removes frozen panes after validation and writes the workbook using XlsbSaveOptions with ValidateMergedAreas set to true.
// Common Searches: Aspose.Cells C# freeze panes at cell E5 and read back frozen rows and columns | How to check if panes are frozen in an Excel workbook with Aspose.Cells .NET | C# Aspose.Cells unfreeze panes before saving the file | Saving an Excel workbook as .xlsb with ValidateMergedAreas using Aspose.Cells | Load existing workbook and apply FreezePanes overloads in Aspose.Cells for .NET
// Tags: Aspose.Cells FreezePanes overloads C# | Aspose.Cells GetFreezedPanes verification | Aspose.Cells UnFreezePanes method | Aspose.Cells Workbook.Save XlsbSaveOptions | Aspose.Cells load workbook apply freeze panes

using System;
using Aspose.Cells;

namespace FreezePaneWorkflowDemo
{
    // // Demonstrates loading or creating a workbook, accessing the first worksheet, applying FreezePanes via index and cell-name overloads, retrieving freeze settings with GetFreezedPanes, optionally unfreezing panes, and saving the file (including optional XlsbSaveOptions with ValidateMergedAreas).
    class Program
    {
        static void Main(string[] args)
        {
            // -------------------------------------------------
            // 1. Create a new workbook (or load an existing one)
            // -------------------------------------------------
            // Creating a fresh workbook instance
            Workbook workbook = new Workbook();

            // If you need to load an existing file, uncomment the line below
            // Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // -------------------------------------------------
            // 2. Access the target worksheet
            // -------------------------------------------------
            Worksheet worksheet = workbook.Worksheets[0]; // First worksheet

            // -------------------------------------------------
            // 3. Freeze panes
            // -------------------------------------------------
            // Example 1: Freeze using row/column indices
            // Freeze at cell C3 (row index 2, column index 2) with 3 rows and 3 columns frozen
            worksheet.FreezePanes(2, 2, 3, 3);

            // Example 2: Freeze using cell name
            // This will overwrite the previous freeze settings
            worksheet.FreezePanes("E5", 4, 2); // Freeze at cell E5 with 4 rows and 2 columns frozen

            // -------------------------------------------------
            // 4. Validate that panes are frozen
            // -------------------------------------------------
            int frozenRow, frozenColumn, frozenRows, frozenColumns;
            bool isFrozen = worksheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);

            Console.WriteLine("Freeze panes applied: " + isFrozen);
            if (isFrozen)
            {
                Console.WriteLine($"Freeze position - Row: {frozenRow}, Column: {frozenColumn}");
                Console.WriteLine($"Number of frozen rows: {frozenRows}, frozen columns: {frozenColumns}");
            }

            // -------------------------------------------------
            // 5. (Optional) Unfreeze panes and verify
            // -------------------------------------------------
            worksheet.UnFreezePanes();

            bool isStillFrozen = worksheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);
            Console.WriteLine("After UnFreezePanes, is frozen: " + isStillFrozen);

            // -------------------------------------------------
            // 6. Save the workbook
            // -------------------------------------------------
            // Save to a new file
            workbook.Save("FreezePaneWorkflowResult.xlsx");

            // If you need to use specific save options (e.g., validate merged areas), you can do:
            // XlsbSaveOptions saveOptions = new XlsbSaveOptions { ValidateMergedAreas = true };
            // workbook.Save("FreezePaneWorkflowResult.xlsb", saveOptions);
        }
    }
}
