// Title: C# Integration Test: Verify FreezePanes Persistence in Aspose.Cells Workbook
// Description: Creates a temporary workbook, applies FreezePanes to columns A‑C, saves the file, reloads it, and uses GetFreezedPanes to assert that the frozen rows and columns remain unchanged, confirming persistence of pane settings after a save‑load cycle.
// Keywords: Aspose.Cells | FreezePanes | GetFreezedPanes | C# integration test | worksheet freeze persistence | .NET Excel automation | save and reload workbook | unit test Aspose.Cells
// Common Searches: Aspose.Cells test frozen columns after save | C# verify FreezePanes persistence | GetFreezedPanes example after reload | integration test for Excel pane freezing | Aspose.Cells FreezePanes unit test
// Developer Intent: Confirm that column freeze settings survive saving and reloading a workbook using Aspose.Cells for .NET.
// Use Cases: Automated CI validation that FreezePanes(0,3) is retained in generated reports. | Regression test for Excel exports that rely on frozen columns for user navigation. | Quality‑gate check in a data pipeline to ensure pane freezing is not lost during file serialization.
// AI Prompts: Generate an MSTest method that creates a workbook, freezes columns A‑C with FreezePanes, saves, reloads, and asserts GetFreezedPanes values. | Write a NUnit test for Aspose.Cells that verifies frozen rows and columns persist after a save‑load operation. | Provide an xUnit example that checks FreezePanes persistence in a temporary Excel file using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsIntegrationTests
{
    // Creates a temporary workbook, applies FreezePanes to columns A‑C, saves the file, reloads it, and uses GetFreezedPanes to assert that the frozen rows and columns remain unchanged, confirming persistence of pane settings after a save‑load cycle.
    class FreezePanesPersistenceTest
    {
        static void Main()
        {
            // Create a temporary file path for the workbook
            string tempFile = Path.Combine(Path.GetTempPath(), "FreezePanesTest.xlsx");

            // ---------- Create workbook and freeze columns ----------
            Workbook workbook = new Workbook();                     // create new workbook
            Worksheet sheet = workbook.Worksheets[0];               // get first worksheet

            // Freeze the first three columns (A, B, C). Row index = 0, column index = 3 (D)
            // frozenRows = 0 (no rows frozen), frozenColumns = 3 (columns A‑C frozen)
            sheet.FreezePanes(0, 3, 0, 3);

            // Save the workbook to the temporary file
            workbook.Save(tempFile);

            // ---------- Load workbook and verify freeze panes ----------
            Workbook loadedWorkbook = new Workbook(tempFile);       // load saved workbook
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0]; // get first worksheet

            // Retrieve freeze pane information
            bool hasFreeze = loadedSheet.GetFreezedPanes(
                out int row, out int column, out int frozenRows, out int frozenColumns);

            // Validate that the freeze settings persisted
            if (!hasFreeze)
                throw new Exception("Freeze panes were not detected after reloading the workbook.");

            if (row != 0 || column != 3)
                throw new Exception($"Unexpected freeze position. Expected row=0, column=3 but got row={row}, column={column}.");

            if (frozenRows != 0 || frozenColumns != 3)
                throw new Exception($"Unexpected frozen size. Expected frozenRows=0, frozenColumns=3 but got frozenRows={frozenRows}, frozenColumns={frozenColumns}.");

            Console.WriteLine("Freeze panes persisted correctly after save and reload.");

            // Clean up temporary file (optional)
            try { File.Delete(tempFile); } catch { /* ignore cleanup errors */ }
        }
    }
}
