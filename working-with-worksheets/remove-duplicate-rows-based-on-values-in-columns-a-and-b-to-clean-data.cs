using System;
using Aspose.Cells;

namespace RemoveDuplicateRowsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the last row that contains data (zero‑based index)
            int lastDataRow = cells.MaxDataRow;

            // Define the range that includes columns A (0) and B (1)
            int startRow = 0;          // include header row
            int startColumn = 0;       // column A
            int endRow = lastDataRow;  // last row with data
            int endColumn = 1;         // column B

            // Specify that the range has headers and that duplicates are checked based on columns A and B
            bool hasHeaders = true;
            int[] columnOffsets = new int[] { 0, 1 }; // offsets for columns A and B

            // Remove duplicate rows based on the values in columns A and B
            cells.RemoveDuplicates(startRow, startColumn, endRow, endColumn, hasHeaders, columnOffsets);

            // Save the cleaned workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}