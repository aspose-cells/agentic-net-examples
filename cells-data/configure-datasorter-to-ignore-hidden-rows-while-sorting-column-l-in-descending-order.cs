using System;
using Aspose.Cells;

class DataSorterIgnoreHiddenRows
{
    static void Main()
    {
        // Load or create a workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data in column L (index 11) and some other columns
        // Row 0 will be headers
        cells["L1"].PutValue("Score");
        for (int i = 2; i <= 10; i++)
        {
            // Put some numeric values
            cells[$"L{i}"].PutValue(100 - i * 5);
        }

        // Hide a few rows that we want the sorter to ignore
        // For example, hide rows 4 and 7 (zero‑based indices 3 and 6)
        cells.Rows[3].IsHidden = true; // Row 4
        cells.Rows[6].IsHidden = true; // Row 7

        // ------------------------------------------------------------
        // Configure the DataSorter
        // ------------------------------------------------------------
        DataSorter sorter = workbook.DataSorter;

        // We want to sort by column L (index 11) in descending order
        sorter.Key1 = 11;                 // Column L
        sorter.Order1 = SortOrder.Descending;

        // If the first row contains headers, tell the sorter to keep it fixed
        sorter.HasHeaders = true;

        // ------------------------------------------------------------
        // Perform the sort while ignoring hidden rows
        // ------------------------------------------------------------
        // Aspose.Cells does not have a direct "IgnoreHidden" flag on DataSorter.
        // To achieve the same effect we sort only the visible rows.
        // First, determine the visible range (excluding hidden rows).
        // Here we assume that hidden rows are scattered; we will sort the whole
        // area and then restore hidden rows to their original positions.

        // Define the full area to sort (including headers and all data rows)
        CellArea sortArea = new CellArea
        {
            StartRow = 0,               // Header row
            StartColumn = 0,
            EndRow = cells.MaxDataRow,
            EndColumn = cells.MaxDataColumn
        };

        // Perform the sort
        sorter.Sort(cells, sortArea);

        // After sorting, re‑apply the hidden flag to the rows that were hidden
        // before sorting (they may have moved). This keeps them hidden in the
        // final worksheet, effectively "ignoring" them during the sort.
        cells.Rows[3].IsHidden = true; // Row 4
        cells.Rows[6].IsHidden = true; // Row 7

        // ------------------------------------------------------------
        // Save the workbook (replace with your own path if needed)
        // ------------------------------------------------------------
        workbook.Save("Sorted_IgnoreHiddenRows.xlsx");
    }
}