using System;
using Aspose.Cells;

namespace AsposeCellsSortingExample
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a DataSorter instance from the workbook
            DataSorter sorter = workbook.DataSorter;

            // Indicate that the first row contains headers (so they are not sorted)
            sorter.HasHeaders = true;

            // Add a sort key for column D (zero‑based index 3) with descending order
            // This preserves the original row grouping because the entire rows are sorted together.
            sorter.AddKey(3, SortOrder.Descending);

            // Define the range to sort:
            // Start at row 0, column 0 (top‑left cell of the data)
            // End at the last used row and column in the worksheet
            int startRow = 0;
            int startColumn = 0;
            int endRow = worksheet.Cells.MaxDataRow;
            int endColumn = worksheet.Cells.MaxDataColumn;

            // Perform the sort on the defined range
            sorter.Sort(worksheet.Cells, startRow, startColumn, endRow, endColumn);

            // Save the sorted workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}