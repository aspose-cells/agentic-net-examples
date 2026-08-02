using System;
using Aspose.Cells;

namespace CustomPrioritySortExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // Replace with new Workbook("input.xlsx") if needed
            Worksheet worksheet = workbook.Worksheets[0];

            // Assume data is already present in the worksheet.
            // Column M (zero‑based index 12) contains priority values: High, Medium, Low.

            // Get the DataSorter object
            DataSorter sorter = workbook.DataSorter;

            // Indicate that the range has a header row
            sorter.HasHeaders = true;

            // Define a custom sort list for the priority column
            // The list order is High → Medium → Low
            string customList = "High,Medium,Low";

            // Add the sort key: column index 12 (M), ascending order, with the custom list
            sorter.AddKey(12, SortOrder.Ascending, customList);

            // Define the range to sort (including all columns with data)
            CellArea sortArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = worksheet.Cells.MaxDataRow,
                EndColumn = worksheet.Cells.MaxDataColumn
            };

            // Perform the sort
            sorter.Sort(worksheet.Cells, sortArea);

            // Save the workbook
            workbook.Save("SortedByPriority.xlsx");
        }
    }
}