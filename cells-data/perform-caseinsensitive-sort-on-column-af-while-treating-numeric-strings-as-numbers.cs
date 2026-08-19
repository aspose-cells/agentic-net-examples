// Title: C# – Case‑Insensitive Sort of Column AF with Numeric Strings as Numbers using Aspose.Cells
// Description: Load an Excel workbook, locate column AF (index 31), define a single‑column sort range, and use Aspose.Cells DataSorter to sort the column case‑insensitively while treating numeric strings as numbers. The header row is preserved and the sorted workbook is saved.
// Keywords: Aspose.Cells | C# | .NET | DataSorter | case insensitive sort | numeric string sorting | Excel column AF | sort with header row | Excel automation | Aspose.Cells example
// Common Searches: Aspose.Cells sort column case insensitive C# | DataSorter treat numeric strings as numbers | How to sort column AF in Excel with Aspose.Cells | Preserve header row while sorting Excel column using .NET | Sort mixed text and numbers in Excel programmatically
// Developer Intent: Sort column AF in an Excel file case‑insensitively, interpret numeric strings as numbers, and keep any header row at the top using Aspose.Cells for .NET.
// Use Cases: Organize a product catalog where codes like "A10" and "a2" need natural numeric ordering without case bias. | Prepare a financial report that mixes case‑variant text and numeric IDs in column AF, ensuring correct ascending order before distribution. | Clean imported data by sorting a single column while retaining the header, then export the workbook for downstream processing.
// AI Prompts: Generate C# code that uses Aspose.Cells DataSorter to sort column AF case‑insensitively, treating numeric strings as numbers and preserving the header row. | Show how to sort multiple columns with different sort directions in Aspose.Cells while keeping headers intact. | Explain the configuration needed in Aspose.Cells to sort a column containing both text and numeric strings without case sensitivity.

using System;
using Aspose.Cells;

// Load an Excel workbook, locate column AF (index 31), define a single‑column sort range, and use Aspose.Cells DataSorter to sort the column case‑insensitively while treating numeric strings as numbers. The header row is preserved and the sorted workbook is saved.
class SortColumnAF
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Column AF -> index 31 (A=0)
        int columnIndex = CellsHelper.ColumnNameToIndex("AF");

        // Determine the last row that contains data
        int lastRow = cells.MaxDataRow;

        // Define the area to sort (entire column AF)
        CellArea sortArea = new CellArea
        {
            StartRow = 0,
            EndRow = lastRow,
            StartColumn = columnIndex,
            EndColumn = columnIndex
        };

        // Configure the DataSorter
        DataSorter sorter = workbook.DataSorter;
        sorter.CaseSensitive = false;      // case‑insensitive
        sorter.SortAsNumber = true;        // treat numeric strings as numbers
        sorter.HasHeaders = true;          // keep header row (if any) at top

        // Add sort key for column AF (ascending)
        sorter.AddKey(columnIndex, SortOrder.Ascending);

        // Perform the sort
        sorter.Sort(cells, sortArea);

        // Save the workbook with the sorted data
        workbook.Save("output.xlsx");
    }
}
