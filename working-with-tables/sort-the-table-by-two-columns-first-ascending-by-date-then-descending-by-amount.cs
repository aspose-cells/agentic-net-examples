using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook that contains the table to be sorted.
        // Replace "input.xlsx" with the actual file path.
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure the DataSorter.
        // Assume the first row contains headers.
        // Column A (index 0) holds the Date values.
        // Column B (index 1) holds the Amount values.
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true;

        // First key: Date column, ascending order.
        sorter.Key1 = 0;               // Date column (A)
        sorter.Order1 = SortOrder.Ascending;

        // Second key: Amount column, descending order.
        sorter.Key2 = 1;               // Amount column (B)
        sorter.Order2 = SortOrder.Descending;

        // Define the area to sort (including headers and all data rows/columns).
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = worksheet.Cells.MaxDataRow,
            EndColumn = worksheet.Cells.MaxDataColumn
        };

        // Perform the sort.
        sorter.Sort(worksheet.Cells, sortArea);

        // Save the sorted workbook.
        // Replace "sorted_output.xlsx" with the desired output path.
        workbook.Save("sorted_output.xlsx");
    }
}