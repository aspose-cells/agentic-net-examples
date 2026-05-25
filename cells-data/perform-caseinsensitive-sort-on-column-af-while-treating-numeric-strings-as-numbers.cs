using System;
using Aspose.Cells;

class SortColumnAF
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Column AF index (A=0, B=1, ..., Z=25, AA=26, ..., AF=31)
        int columnIndex = CellsHelper.ColumnNameToIndex("AF");

        // Determine the range that contains data (including all columns)
        int startRow = 0; // assuming the first row contains headers
        int endRow = cells.MaxDataRow;
        int startColumn = 0;
        int endColumn = cells.MaxDataColumn;

        // Define the area to be sorted
        CellArea sortArea = new CellArea
        {
            StartRow = startRow,
            EndRow = endRow,
            StartColumn = startColumn,
            EndColumn = endColumn
        };

        // Configure the DataSorter
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders = true;          // keep header row at top
        sorter.CaseSensitive = false;      // case‑insensitive sorting
        sorter.SortAsNumber = true;        // treat numeric strings as numbers
        sorter.AddKey(columnIndex, SortOrder.Ascending); // sort by column AF

        // Perform the sort
        sorter.Sort(cells, sortArea);

        // Save the sorted workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}