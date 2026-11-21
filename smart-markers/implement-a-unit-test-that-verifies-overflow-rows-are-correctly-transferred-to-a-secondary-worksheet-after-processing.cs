// Title: C# unit test to verify that rows exceeding a maximum per sheet are transferred to a secondary worksheet using Aspose.Cells
// Common Searches: aspocells unit test for row overflow handling in C# | how to move rows beyond max rows per sheet to a new worksheet with Aspose.Cells | verify overflow rows are copied to secondary sheet in Aspose.Cells C# example | C# Aspose.Cells test for transferring excess rows to another worksheet | unit testing Aspose.Cells TransferOverflowRows method in .NET
// Tags: aspocells transfer overflow rows | c# aspocells secondary worksheet creation | aspocells max rows per sheet handling | unit test aspocells row overflow | aspocells workbook row deletion verification

using System;
using Aspose.Cells;

namespace AsposeCellsOverflowTests
{
    // The example creates a workbook with 150 rows, invokes a TransferOverflowRows method that moves rows beyond the defined limit of 100 to a newly added worksheet named "OverflowRows", and then runs assertions to confirm the primary sheet contains exactly 100 rows with correct data, the secondary sheet exists with the remaining 50 rows, and the transferred rows retain their original values. Any discrepancy throws an exception; otherwise a success message is displayed.
    class Program
    {
        // Maximum number of rows allowed per worksheet before overflow handling
        private const int MaxRowsPerSheet = 100;

        // Transfers rows exceeding maxRows from the primary worksheet to a new secondary worksheet.
        private static void TransferOverflowRows(Workbook workbook, int maxRows)
        {
            // Assume the primary worksheet is the first one
            Worksheet primarySheet = workbook.Worksheets[0];
            Cells primaryCells = primarySheet.Cells;

            // Determine the last used row in the primary sheet (zero‑based)
            int lastUsedRow = primaryCells.MaxDataRow;

            // No overflow rows – nothing to do
            if (lastUsedRow < maxRows) return;

            // Create secondary worksheet
            int newIndex = workbook.Worksheets.Add();
            Worksheet secondarySheet = workbook.Worksheets[newIndex];
            secondarySheet.Name = "OverflowRows";
            Cells secondaryCells = secondarySheet.Cells;

            // Copy overflow rows to secondary sheet
            for (int srcRow = maxRows; srcRow <= lastUsedRow; srcRow++)
            {
                int destRow = srcRow - maxRows; // start at row 0 in secondary sheet
                for (int col = 0; col <= primaryCells.MaxDataColumn; col++)
                {
                    Cell srcCell = primaryCells[srcRow, col];
                    if (srcCell != null && srcCell.Value != null)
                    {
                        secondaryCells[destRow, col].Value = srcCell.Value;
                    }
                }
            }

            // Remove overflow rows from primary sheet
            // Deleting a row shifts subsequent rows up, so we stay on the same index.
            for (int row = maxRows; row <= lastUsedRow; row++)
            {
                primaryCells.DeleteRow(row);
                row--;
                lastUsedRow--;
            }
        }

        static void Main()
        {
            try
            {
                // ---------- Create ----------
                // Create a new workbook with a single worksheet
                Workbook workbook = new Workbook();
                Worksheet primarySheet = workbook.Worksheets[0];
                Cells cells = primarySheet.Cells;

                // Populate the primary worksheet with 150 rows of sample data
                for (int row = 0; row < 150; row++)
                {
                    cells[row, 0].Value = $"Row {row + 1}";
                    cells[row, 1].Value = row + 1; // numeric value
                }

                // ---------- Process ----------
                // Transfer rows beyond MaxRowsPerSheet to a secondary worksheet
                TransferOverflowRows(workbook, MaxRowsPerSheet);

                // ---------- Verify ----------
                // Primary sheet should now contain exactly MaxRowsPerSheet rows with data
                Worksheet resultPrimary = workbook.Worksheets[0];
                Cells primaryResultCells = resultPrimary.Cells;
                int primaryLastRow = primaryResultCells.MaxDataRow; // zero‑based

                if (primaryLastRow != MaxRowsPerSheet - 1)
                    throw new Exception($"Primary sheet last row index {primaryLastRow} does not match expected {MaxRowsPerSheet - 1}.");

                // Verify content of the last row in primary sheet (should be row 100)
                if (primaryResultCells[MaxRowsPerSheet - 1, 0].StringValue != "Row 100")
                    throw new Exception("Primary sheet last row text mismatch.");
                if (primaryResultCells[MaxRowsPerSheet - 1, 1].IntValue != 100)
                    throw new Exception("Primary sheet last row numeric value mismatch.");

                // Secondary sheet should exist and contain the overflow rows (rows 101‑150)
                Worksheet secondarySheet = workbook.Worksheets["OverflowRows"];
                if (secondarySheet == null)
                    throw new Exception("Secondary worksheet for overflow rows was not created.");

                Cells secondaryCells = secondarySheet.Cells;
                int overflowRowCount = secondaryCells.MaxDataRow + 1; // convert zero‑based to count

                if (overflowRowCount != 150 - MaxRowsPerSheet)
                    throw new Exception($"Overflow row count {overflowRowCount} does not match expected {150 - MaxRowsPerSheet}.");

                // Verify first overflow row (original row 101) now appears as row 0 in secondary sheet
                if (secondaryCells[0, 0].StringValue != "Row 101")
                    throw new Exception("First overflow row text mismatch.");
                if (secondaryCells[0, 1].IntValue != 101)
                    throw new Exception("First overflow row numeric value mismatch.");

                // Verify last overflow row (original row 150) now appears as last row in secondary sheet
                int lastOverflowIndex = overflowRowCount - 1;
                if (secondaryCells[lastOverflowIndex, 0].StringValue != "Row 150")
                    throw new Exception("Last overflow row text mismatch.");
                if (secondaryCells[lastOverflowIndex, 1].IntValue != 150)
                    throw new Exception("Last overflow row numeric value mismatch.");

                Console.WriteLine("All verifications passed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }
    }
}
