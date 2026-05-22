using System;
using Aspose.Cells;

namespace AsposeCellsDuplicateValidation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the used range of the worksheet
            int startRow = 0;                         // Assuming the first row contains headers
            int startColumn = 0;
            int endRow = cells.MaxDataRow;            // Last row with data
            int endColumn = cells.MaxDataColumn;      // Last column with data

            // Record the original number of data rows (including header)
            int originalRowCount = endRow - startRow + 1;

            // Define which columns constitute the key for duplicate detection.
            // Example: first two columns (A and B) are the key columns.
            int[] keyColumnOffsets = new int[] { 0, 1 };

            // Perform duplicate removal on a copy of the range to test for duplicates.
            // The method returns void, so we compare row counts before and after.
            // hasHeaders = true because the first row is a header row.
            cells.RemoveDuplicates(startRow + 1, startColumn, endRow, endColumn, true, keyColumnOffsets);

            // After removal, recalculate the last data row.
            int newEndRow = cells.MaxDataRow;
            int newRowCount = newEndRow - startRow + 1;

            // Validate: if row counts are equal, no duplicates existed; otherwise duplicates were present.
            if (newRowCount == originalRowCount)
            {
                Console.WriteLine("Validation passed: No duplicate rows based on the key columns.");
            }
            else
            {
                Console.WriteLine($"Validation failed: {originalRowCount - newRowCount} duplicate row(s) were found and removed.");
            }

            // Save the workbook (optional – the file now contains the deduplicated data)
            workbook.Save("output.xlsx");
        }
    }
}