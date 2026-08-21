// Title: C# – Validate and Remove Duplicate Rows in an Aspose.Cells Worksheet Using Key Columns
// Description: This example creates a workbook, populates it with sample data, and demonstrates how to use Aspose.Cells' RemoveDuplicates method to detect and delete rows that share the same values in the ID and Name columns. It reports the row count before and after removal and saves the cleaned file.
// Keywords: Aspose.Cells RemoveDuplicates C# | duplicate row validation .NET | Excel duplicate detection Aspose | key column duplicate removal | C# Aspose.Cells example
// Common Searches: remove duplicate rows Aspose.Cells C# | validate Excel table for duplicate key columns | Aspose.Cells RemoveDuplicates multiple columns | C# code to delete duplicate rows in Excel | how to check for duplicate rows with Aspose.Cells
// Developer Intent: Identify rows that have identical values in specified key columns and eliminate those duplicates programmatically.
// Use Cases: Clean imported CSV data before analysis by ensuring unique ID‑Name pairs. | Maintain a master employee list without repeated records. | Prevent duplicate transaction entries in financial worksheets. | Prepare data for reporting pipelines that require unique key combinations.
// AI Prompts: Generate C# code that lists duplicate rows in an Aspose.Cells worksheet without removing them. | Show how to log each removed row’s original index and values when using RemoveDuplicates. | Provide an alternative implementation using a HashSet to flag duplicate key combinations in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsDuplicateValidation
{
    // This example creates a workbook, populates it with sample data, and demonstrates how to use Aspose.Cells' RemoveDuplicates method to detect and delete rows that share the same values in the ID and Name columns. It reports the row count before and after removal and saves the cleaned file.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ---------- Populate sample data ----------
            // Header row
            cells[0, 0].PutValue("ID");
            cells[0, 1].PutValue("Name");
            cells[0, 2].PutValue("Age");

            // Data rows (some duplicates based on ID + Name)
            cells[1, 0].PutValue(1); cells[1, 1].PutValue("John");   cells[1, 2].PutValue(30);
            cells[2, 0].PutValue(2); cells[2, 1].PutValue("Jane");   cells[2, 2].PutValue(25);
            cells[3, 0].PutValue(1); cells[3, 1].PutValue("John");   cells[3, 2].PutValue(31); // duplicate key
            cells[4, 0].PutValue(3); cells[4, 1].PutValue("Bob");    cells[4, 2].PutValue(40);
            cells[5, 0].PutValue(2); cells[5, 1].PutValue("Jane");   cells[5, 2].PutValue(26); // duplicate key
            cells[6, 0].PutValue(4); cells[6, 1].PutValue("Alice");  cells[6, 2].PutValue(22);

            // ---------- Determine the data range ----------
            int startRow = 1;                     // first data row (skip header)
            int startColumn = 0;                  // first column (ID)
            int endRow = cells.MaxDataRow;        // last row with data
            int endColumn = cells.MaxDataColumn;  // last column with data

            // Record original number of data rows
            int originalRowCount = endRow - startRow + 1;

            // ---------- Validate duplicates using RemoveDuplicates ----------
            // Key columns are ID (offset 0) and Name (offset 1)
            int[] keyColumnOffsets = new int[] { 0, 1 };
            // The method removes duplicate rows in-place
            cells.RemoveDuplicates(startRow, startColumn, endRow, endColumn, true, keyColumnOffsets);

            // After removal, recalculate the last data row
            int newEndRow = cells.MaxDataRow;
            int newRowCount = newEndRow - startRow + 1;

            // ---------- Output validation result ----------
            if (newRowCount < originalRowCount)
            {
                Console.WriteLine("Duplicates were found and removed.");
                Console.WriteLine($"Rows before: {originalRowCount}, rows after: {newRowCount}");
            }
            else
            {
                Console.WriteLine("No duplicate rows based on the specified key columns were found.");
            }

            // ---------- Save the workbook ----------
            workbook.Save("DuplicateValidationResult.xlsx");
        }
    }
}
