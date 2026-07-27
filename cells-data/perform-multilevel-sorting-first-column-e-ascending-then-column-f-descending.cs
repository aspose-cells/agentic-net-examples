using System;
using Aspose.Cells;

class MultiLevelSortExample
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the DataSorter object associated with the workbook
        DataSorter sorter = workbook.DataSorter;

        // The data has a header row
        sorter.HasHeaders = true;

        // First sort key: column E (index 4) – ascending
        sorter.Key1 = 4;                     // Column E
        sorter.Order1 = SortOrder.Ascending;

        // Second sort key: column F (index 5) – descending
        sorter.Key2 = 5;                     // Column F
        sorter.Order2 = SortOrder.Descending;

        // Define the area to be sorted (including the header row)
        CellArea sortArea = new CellArea
        {
            StartRow = 0,                                 // Header row
            StartColumn = 4,                              // Column E
            EndRow = worksheet.Cells.MaxDataRow,          // Last used row
            EndColumn = 5                                 // Column F
        };

        // Perform the multi‑level sort
        sorter.Sort(worksheet.Cells, sortArea);

        // Save the sorted workbook (replace with your desired output path)
        workbook.Save("sorted_output.xlsx");
    }
}