using System;
using Aspose.Cells;

namespace AsposeCellsSortingExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // Initialize the DataSorter
            DataSorter sorter = workbook.DataSorter;

            // Assume the first row contains headers
            sorter.HasHeaders = true;

            // Preserve original grouping by sorting first on the grouping column (e.g., column C, index 2)
            sorter.AddKey(2, SortOrder.Ascending);

            // Then sort column D (index 3) in descending order
            sorter.AddKey(3, SortOrder.Descending);

            // Define the range to sort (adjust the end row/column as needed)
            // Here we sort from A1 to D (last used row) to include all relevant data
            int lastRow = worksheet.Cells.MaxDataRow;
            CellArea sortArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = lastRow,
                EndColumn = 3 // Column D index
            };

            // Perform the sort
            sorter.Sort(worksheet.Cells, sortArea);

            // Save the sorted workbook
            workbook.Save("output.xlsx");
        }
    }
}