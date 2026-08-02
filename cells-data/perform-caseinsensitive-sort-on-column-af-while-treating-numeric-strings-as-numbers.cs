// Title: Sort Column AF Case‑Insensitive with Numeric Strings as Numbers – Aspose.Cells C#
// Description: Loads an Excel workbook, defines the used rows, and sorts column AF (index 31) in ascending order. The Aspose.Cells DataSorter is configured to keep the header row, ignore case, and treat numeric strings as true numbers before saving the result.
// Keywords: Aspose.Cells sort column C# | DataSorter case insensitive | numeric string sorting Aspose | Excel column AF sort | sort with headers Aspose.Cells | C# Excel sorting example
// Common Searches: Aspose.Cells sort a single column case insensitive | Treat numeric strings as numbers when sorting Excel with Aspose | How to keep header row while sorting column AF in C# | Sort column AF ascending using Aspose.Cells DataSorter
// Developer Intent: Perform a case‑insensitive ascending sort on column AF while converting numeric‑like strings to numbers, preserving the header row.
// Use Cases: Arrange product codes that mix letters and numbers, ensuring "10" follows "2". | Order employee IDs with leading zeros so they sort numerically rather than lexicographically. | Prepare a report where column AF must be alphabetically ordered without case bias, with numeric strings sorted by value.
// AI Prompts: Write C# code using Aspose.Cells to sort column AF case‑insensitively and treat numeric strings as numbers, keeping the first row as a header. | Explain the effect of DataSorter properties HasHeaders, CaseSensitive, and SortAsNumber in Aspose.Cells. | Show how to sort multiple columns with different directions using Aspose.Cells DataSorter in C#.

using System;
using Aspose.Cells;

// Loads an Excel workbook, defines the used rows, and sorts column AF (index 31) in ascending order. The Aspose.Cells DataSorter is configured to keep the header row, ignore case, and treat numeric strings as true numbers before saving the result.
class SortColumnAF
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Determine the used range of rows
        int startRow = cells.MinDataRow;          // usually 0 (including header)
        int endRow   = cells.MaxDataRow;          // last row with data

        // Column AF is the 32nd column (zero‑based index 31)
        const int columnAFIndex = 31;

        // Define the area to sort: all rows in column AF (including header if present)
        CellArea sortArea = new CellArea
        {
            StartRow = startRow,
            EndRow   = endRow,
            StartColumn = columnAFIndex,
            EndColumn   = columnAFIndex
        };

        // Configure the DataSorter
        DataSorter sorter = workbook.DataSorter;
        sorter.HasHeaders   = true;      // assume first row is a header
        sorter.CaseSensitive = false;    // case‑insensitive sorting
        sorter.SortAsNumber = true;      // treat numeric strings as numbers

        // Add the sort key for column AF (ascending order)
        sorter.AddKey(columnAFIndex, SortOrder.Ascending);

        // Perform the sort
        sorter.Sort(cells, sortArea);

        // Save the sorted workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
