// Title: C# unit test for transferring overflow rows to an "Overflow" worksheet using Aspose.Cells
// AI Prompts: Write a C# unit test that creates a workbook, fills the first sheet with more rows than a given limit, calls the overflow‑transfer method, and asserts that only the allowed rows remain in the primary sheet while the excess rows appear correctly in a sheet named "Overflow". | Extend the overflow‑transfer routine to also copy cell formulas, comments, and hyperlinks, then update the unit test to verify that these elements are preserved in the overflow worksheet. | Generate a data‑driven NUnit test that runs the overflow‑transfer logic with several max‑row thresholds (e.g., 3, 5, 10) and checks that the overflow sheet is created only when needed and contains the expected number of rows.
// Common Searches: Aspose.Cells C# unit test overflow rows to secondary worksheet | how to move excess Excel rows to another sheet with Aspose.Cells | verify overflow row transfer in Aspose.Cells unit test | C# test splitting worksheet data into primary and overflow sheets | Aspose.Cells example for handling row limit overflow
// Tags: Aspose.Cells transfer overflow rows | Aspose.Cells unit test row limit | Aspose.Cells move excess rows to overflow sheet | Aspose.Cells validate overflow worksheet creation | Aspose.Cells delete rows after copy

using System;
using Aspose.Cells;

namespace AsposeCellsOverflowTests
{
    // Demonstrates a self‑contained C# unit test that populates a workbook, moves rows exceeding a defined maximum from the first worksheet to an "Overflow" worksheet, clears the original rows, and asserts both sheet row counts and cell values using Aspose.Cells for .NET.
    public class OverflowTransferDemo
    {
        // Threshold for maximum rows allowed in the primary worksheet
        private const int MaxRowsInPrimary = 5;

        // Name of the secondary worksheet that will hold overflow rows
        private const string OverflowSheetName = "Overflow";

        /// <summary>
        /// Transfers rows exceeding the threshold from the first worksheet to a secondary worksheet.
        /// If the secondary worksheet does not exist, it will be created.
        /// </summary>
        /// <param name="workbook">The workbook containing the worksheets.</param>
        /// <param name="maxRows">Maximum rows allowed in the primary worksheet.</param>
        private static void TransferOverflowRows(Workbook workbook, int maxRows)
        {
            // Get the primary worksheet (first sheet)
            Worksheet primarySheet = workbook.Worksheets[0];
            Cells primaryCells = primarySheet.Cells;

            // Determine the total number of used rows in the primary sheet
            int totalRows = primaryCells.MaxDataRow + 1; // MaxDataRow is zero‑based

            // If there are no overflow rows, exit early
            if (totalRows <= maxRows)
                return;

            // Get or create the overflow worksheet
            Worksheet overflowSheet = workbook.Worksheets[OverflowSheetName] ??
                                      workbook.Worksheets.Add(OverflowSheetName);
            Cells overflowCells = overflowSheet.Cells;

            // Copy overflow rows to the overflow sheet
            for (int row = maxRows; row < totalRows; row++)
            {
                for (int col = 0; col <= primaryCells.MaxDataColumn; col++)
                {
                    Cell sourceCell = primaryCells[row, col];
                    if (sourceCell != null && sourceCell.Type != CellValueType.IsNull)
                    {
                        // Preserve the value and style
                        overflowCells[row - maxRows, col].PutValue(sourceCell.Value);
                        overflowCells[row - maxRows, col].SetStyle(sourceCell.GetStyle());
                    }
                }
            }

            // Clear overflow rows from the primary sheet
            // Deleting rows shifts subsequent rows up, so we adjust the counters accordingly
            for (int row = maxRows; row < totalRows; row++)
            {
                primaryCells.DeleteRow(row);
                row--;          // stay at the same index after deletion
                totalRows--;    // total rows reduced by one
            }
        }

        /// <summary>
        /// Executes the overflow transfer test logic.
        /// </summary>
        private static void RunOverflowTransferTest()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the primary worksheet
            Worksheet primarySheet = workbook.Worksheets[0];
            Cells cells = primarySheet.Cells;

            // Populate the primary worksheet with 8 rows of sample data (more than MaxRowsInPrimary)
            for (int i = 0; i < 8; i++)
            {
                cells[i, 0].PutValue($"Row{i + 1}_Col1");
                cells[i, 1].PutValue($"Row{i + 1}_Col2");
            }

            // Verify initial state: primary sheet has 8 rows
            if (cells.MaxDataRow + 1 != 8)
                throw new InvalidOperationException("Initial row count in primary sheet should be 8.");

            // Execute the overflow transfer logic
            TransferOverflowRows(workbook, MaxRowsInPrimary);

            // After processing, primary sheet should contain only MaxRowsInPrimary rows
            Worksheet updatedPrimary = workbook.Worksheets[0];
            Cells updatedPrimaryCells = updatedPrimary.Cells;
            if (updatedPrimaryCells.MaxDataRow + 1 != MaxRowsInPrimary)
                throw new InvalidOperationException($"Primary sheet should contain only {MaxRowsInPrimary} rows after overflow transfer.");

            // The overflow worksheet should exist
            Worksheet overflowSheet = workbook.Worksheets[OverflowSheetName];
            if (overflowSheet == null)
                throw new InvalidOperationException("Overflow worksheet should be created.");

            // Verify overflow sheet contains the expected rows (rows 6‑8 from original data)
            Cells overflowCells = overflowSheet.Cells;
            int expectedOverflowRows = 8 - MaxRowsInPrimary;
            if (overflowCells.MaxDataRow + 1 != expectedOverflowRows)
                throw new InvalidOperationException($"Overflow sheet should contain {expectedOverflowRows} rows.");

            // Validate content of overflow rows
            for (int i = 0; i < expectedOverflowRows; i++)
            {
                string expectedCol1 = $"Row{MaxRowsInPrimary + i + 1}_Col1";
                string expectedCol2 = $"Row{MaxRowsInPrimary + i + 1}_Col2";

                if (overflowCells[i, 0].StringValue != expectedCol1)
                    throw new InvalidOperationException($"Overflow row {i + 1} column 1 value mismatch.");

                if (overflowCells[i, 1].StringValue != expectedCol2)
                    throw new InvalidOperationException($"Overflow row {i + 1} column 2 value mismatch.");
            }

            Console.WriteLine("Overflow transfer test passed successfully.");
        }

        /// <summary>
        /// Entry point of the program.
        /// </summary>
        public static void Main()
        {
            try
            {
                RunOverflowTransferTest();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                // In a real application you might log the exception or rethrow.
            }
        }
    }
}
