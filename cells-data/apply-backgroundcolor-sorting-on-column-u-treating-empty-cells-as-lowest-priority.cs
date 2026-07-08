using System;
using System.Drawing;
using Aspose.Cells;

class BackgroundColorSortExample
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a DataSorter instance
        DataSorter sorter = workbook.DataSorter;

        // Assume the first row contains headers
        sorter.HasHeaders = true;

        // Column U is the 21st column (zero‑based index 20)
        int columnUIndex = 20;

        // Add a sort key that sorts by cell background color in ascending order.
        // Empty cells (no fill) will be treated as the lowest priority.
        sorter.AddKey(columnUIndex, SortOnType.CellColor, SortOrder.Ascending, null);

        // Define the range to sort: from the first row to the last used row in column U
        int lastRow = worksheet.Cells.MaxDataRow;
        CellArea sortArea = CellArea.CreateCellArea($"U1", $"U{lastRow + 1}");

        // Perform the sort
        sorter.Sort(worksheet.Cells, sortArea);

        // Save the sorted workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}