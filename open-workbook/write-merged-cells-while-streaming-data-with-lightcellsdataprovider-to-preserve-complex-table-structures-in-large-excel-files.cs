// Title: Stream large Excel files with LightCellsDataProvider while preserving a merged header – Aspose.Cells C# example
// Description: Demonstrates a custom LightCellsDataProvider that streams 10,000 rows × 5 columns to the first worksheet, merges the entire first row as a title, writes the header value only once, and optionally gathers header strings. The approach creates a memory‑efficient XLSX workbook in .NET without loading the full dataset into RAM.
// Keywords: Aspose.Cells | LightCellsDataProvider | merged cells | streaming large Excel | C# | memory efficient XLSX | large report generation | Excel header merge | .NET
// Common Searches: How to use LightCellsDataProvider to write merged cells in Aspose.Cells | Streaming large Excel workbook with a merged title row C# | Preserve merged header when saving a big XLSX with Aspose.Cells | Aspose.Cells example for memory‑efficient large report | LightCellsDataProvider merge first row across columns
// Developer Intent: Generate a massive Excel workbook by streaming rows and columns with LightCellsDataProvider while keeping a merged header row for visual consistency.
// Use Cases: Create a multi‑megabyte report where the title spans the whole first row without loading all data into memory. | Reduce RAM usage in server‑side reporting services that output XLSX files with thousands of rows. | Apply a global string pool only to header text to minimize file size during streaming. | Integrate fast, low‑memory Excel export into .NET web APIs or background jobs.
// AI Prompts: Show C# code that streams 500,000 rows using LightCellsDataProvider and merges the first row across all columns. | Explain how to add bold formatting and a background color to the merged header cells in the streaming provider. | Provide a way to set column widths while using LightCellsDataProvider for row‑by‑row export. | Give a step‑by‑step guide to modify the example for multiple worksheets, each with its own merged header.

using System;
using Aspose.Cells;

namespace LightCellsMergedExample
{
    // Custom LightCellsDataProvider that streams a large table while preserving merged header cells
    // Demonstrates a custom LightCellsDataProvider that streams 10,000 rows × 5 columns to the first worksheet, merges the entire first row as a title, writes the header value only once, and optionally gathers header strings. The approach creates a memory‑efficient XLSX workbook in .NET without loading the full dataset into RAM.
    public class StreamingDataProvider : LightCellsDataProvider
    {
        private const int MaxRows = 10000;   // total rows to generate (including header)
        private const int MaxCols = 5;       // total columns

        private int _currentRow = -1;
        private int _currentCol = -1;

        // Called for each worksheet; we only handle the first sheet (index 0)
        public bool StartSheet(int sheetIndex)
        {
            return sheetIndex == 0;
        }

        // Return the next row index to be saved; -1 signals no more rows
        public int NextRow()
        {
            _currentRow++;
            _currentCol = -1; // reset column index for the new row
            return _currentRow < MaxRows ? _currentRow : -1;
        }

        // Initialize the row if needed (e.g., set height)
        public void StartRow(Row row)
        {
            // Example: set a larger height for the header row
            if (_currentRow == 0)
                row.Height = 30;
        }

        // Return the next column index for the current row; -1 signals no more cells in this row
        public int NextCell()
        {
            _currentCol++;
            return _currentCol < MaxCols ? _currentCol : -1;
        }

        // Fill the cell with appropriate data
        public void StartCell(Cell cell)
        {
            // Header row (row 0) – merged across all columns, value set only in the first cell
            if (_currentRow == 0 && _currentCol == 0)
            {
                cell.PutValue("Large Report Header");
            }
            // Header titles for remaining columns (optional, can be left blank)
            else if (_currentRow == 0)
            {
                // Leave other cells empty; merged area will display the header value
                cell.PutValue(string.Empty);
            }
            // Data rows
            else
            {
                // Example data: "R{row}C{col}"
                cell.PutValue($"R{_currentRow}C{_currentCol}");
            }
        }

        // Determines whether string values should be gathered into a global pool
        public bool IsGatherString()
        {
            // Gather strings only for the header row to reduce duplication
            return _currentRow == 0;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define a merged header that spans the first row across all columns
            // Merge from row 0, column 0, covering 1 row and MaxCols columns
            cells.Merge(0, 0, 1, 5);

            // Set up save options with the custom LightCellsDataProvider
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new StreamingDataProvider()
            };

            // Save the workbook; the provider streams cell data while preserving the merged header
            workbook.Save("LargeReportWithMergedHeader.xlsx", saveOptions);
        }
    }
}
