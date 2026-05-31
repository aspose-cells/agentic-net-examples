using System;
using Aspose.Cells;

namespace MultiLevelSortingExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the DataSorter object associated with the workbook
            DataSorter sorter = workbook.DataSorter;

            // Assume the first row contains headers
            sorter.HasHeaders = true;

            // Set the first sort key: column E (index 4) ascending
            sorter.Key1 = 4; // Column E
            sorter.Order1 = SortOrder.Ascending;

            // Set the second sort key: column F (index 5) descending
            sorter.Key2 = 5; // Column F
            sorter.Order2 = SortOrder.Descending;

            // Define the range to sort (including headers)
            CellArea sortArea = new CellArea
            {
                StartRow = 0,                     // First row (header)
                StartColumn = 4,                  // Column E
                EndRow = worksheet.Cells.MaxDataRow, // Last row with data
                EndColumn = 5                     // Column F
            };

            // Perform the multi‑level sort
            sorter.Sort(worksheet.Cells, sortArea);

            // Save the sorted workbook (replace with your desired output path)
            workbook.Save("sorted_output.xlsx");
        }
    }
}