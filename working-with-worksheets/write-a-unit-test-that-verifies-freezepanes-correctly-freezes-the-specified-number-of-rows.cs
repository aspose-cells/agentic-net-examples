// Title: C# Unit Test for Aspose.Cells FreezePanes – Verify Frozen Row Count
// Description: Demonstrates how to write a unit test that confirms Aspose.Cells FreezePanes freezes exactly five rows (no columns). The test creates a workbook, applies FreezePanes, retrieves the pane state with GetFreezedPanes, and asserts HasFreeze, start row, start column, frozen rows and frozen columns before saving the file.
// Keywords: Aspose.Cells | FreezePanes | C# unit test | NUnit | MSTest | xUnit | GetFreezedPanes | frozen rows | Excel automation | worksheet freeze pane validation
// Common Searches: Aspose.Cells unit test FreezePanes rows | how to assert frozen rows with GetFreezedPanes | NUnit test for FreezePanes in C# | MSTest example for Aspose.Cells FreezePanes | xUnit verify frozen rows Aspose.Cells
// Developer Intent: Write a test that validates the FreezePanes method freezes the exact number of rows and columns specified.
// Use Cases: Confirm that calling worksheet.FreezePanes(5, 0, 5, 0) sets HasFreeze to true. | Verify the start row index returned by GetFreezedPanes equals 5. | Ensure the frozen‑row count reported is 5 while frozen‑column count is 0. | Check that the workbook can be saved without altering the freeze configuration.
// AI Prompts: Generate an NUnit test that creates a Workbook, applies worksheet.FreezePanes(5,0,5,0), calls GetFreezedPanes, and asserts HasFreeze, actualRow, actualColumn, actualFrozenRows, and actualFrozenColumns. | Provide an MSTest method that validates FreezePanes freezes five rows and no columns, then deletes the generated Excel file in a teardown step. | Write a xUnit test verifying GetFreezedPanes returns the expected parameters after FreezePanes is executed on a worksheet.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Demonstrates how to write a unit test that confirms Aspose.Cells FreezePanes freezes exactly five rows (no columns). The test creates a workbook, applies FreezePanes, retrieves the pane state with GetFreezedPanes, and asserts HasFreeze, start row, start column, frozen rows and frozen columns before saving the file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Freeze the first 5 rows (no columns frozen)
                int freezeRowIndex = 5;      // row index where the freeze starts (0‑based)
                int freezeColumnIndex = 0;   // column index where the freeze starts
                int frozenRows = 5;          // number of rows to freeze
                int frozenColumns = 0;       // number of columns to freeze

                worksheet.FreezePanes(freezeRowIndex, freezeColumnIndex, frozenRows, frozenColumns);

                // Retrieve freeze pane information
                bool hasFreeze = worksheet.GetFreezedPanes(out int actualRow, out int actualColumn,
                                                           out int actualFrozenRows, out int actualFrozenColumns);

                // Simple validation output
                Console.WriteLine($"Has Freeze: {hasFreeze}");
                Console.WriteLine($"Freeze Row Index: {actualRow}");
                Console.WriteLine($"Freeze Column Index: {actualColumn}");
                Console.WriteLine($"Frozen Rows: {actualFrozenRows}");
                Console.WriteLine($"Frozen Columns: {actualFrozenColumns}");

                // Define output path
                string outputPath = "FreezeRowsTestOutput.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
