using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or specify the required one)
        Worksheet worksheet = workbook.Worksheets[0];

        // Determine the last row and column that contain data
        int lastRow = worksheet.Cells.MaxDataRow;
        int lastColumn = worksheet.Cells.MaxDataColumn;

        // Configure the DataSorter to sort by column D (index 3, zero‑based)
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true;               // assume the first row contains headers
        sorter.Key1 = 3;                         // column D
        sorter.Order1 = SortOrder.Ascending;    // sort in ascending order

        // Define the range to be sorted (entire used area)
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = lastRow,
            EndColumn = lastColumn
        };

        // Perform the sort operation
        sorter.Sort(worksheet.Cells, sortArea);

        // Save the workbook (the sorted data will be present when the file is opened)
        workbook.Save("output.xlsx");
    }
}