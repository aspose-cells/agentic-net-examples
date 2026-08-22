// Title: How to sort Excel column AF case‑insensitively as numbers using Aspose.Cells for .NET
// AI Prompts: Apply Aspose.Cells DataSorter to sort column AF without case sensitivity and interpret numeric strings as numeric values. | Create a descending sort on column AF while keeping the header row unchanged, using case‑insensitive numeric sorting with Aspose.Cells. | Modify the sample to sort additional columns, preserving the case‑insensitive numeric sort for column AF.
// Common Searches: asp.net sort column AF case insensitive numeric strings Aspose.Cells | c# Aspose.Cells DataSorter sort single column as number ignoring case | how to sort Excel column AF by numeric value while ignoring case with Aspose.Cells .NET
// Tags: Aspose.Cells DataSorter case-insensitive sorting | numeric string sorting Aspose.Cells | Excel column AF sort .NET | DataSorter specific column range | treat text as numbers Aspose.Cells

using System;
using Aspose.Cells;

// The code loads an Excel workbook, configures Aspose.Cells DataSorter to be case‑insensitive and to treat numeric strings as numbers, adds an ascending sort key for column AF, defines a sort area covering all used rows in that column, executes the sort, and saves the sorted workbook.
class SortColumnAF
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Create a DataSorter instance
        DataSorter sorter = workbook.DataSorter;

        // Set case-insensitive sorting
        sorter.CaseSensitive = false;

        // Treat numeric strings as numbers during sorting
        sorter.SortAsNumber = true;

        // Column AF is the 32nd column (zero‑based index 31)
        int columnIndex = CellsHelper.ColumnNameToIndex("AF");

        // Add sort key for column AF in ascending order
        sorter.AddKey(columnIndex, SortOrder.Ascending);

        // Define the range to sort: from the first data row to the last used row in column AF
        int firstRow = cells.MinDataRow;               // usually 0 (including header if present)
        int lastRow = cells.MaxDataRow;                // last row with data
        CellArea sortArea = new CellArea
        {
            StartRow = firstRow,
            EndRow = lastRow,
            StartColumn = columnIndex,
            EndColumn = columnIndex
        };

        // Perform the sort
        sorter.Sort(cells, sortArea);

        // Save the workbook with the sorted data
        workbook.Save("output.xlsx");
    }
}
