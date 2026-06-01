using System;
using Aspose.Cells;

namespace LightCellsMergedCellsDemo
{
    // Custom LightCellsDataProvider that streams a large table of data.
    // The worksheet already contains merged cells (e.g., a header spanning multiple columns).
    // The provider only supplies cell values; merging is handled beforehand.
    public class LargeTableDataProvider : LightCellsDataProvider
    {
        private const int TotalRows = 1000;      // Number of data rows (excluding header)
        private const int TotalColumns = 5;     // Number of columns in the table
        private int _currentRow = -1;
        private int _currentColumn = -1;

        // Only process the first worksheet (index 0).
        public bool StartSheet(int sheetIndex)
        {
            return sheetIndex == 0;
        }

        // Return the next row index to be saved.
        // Row 0 is the header row (already merged), rows 1..TotalRows contain data.
        public int NextRow()
        {
            _currentRow++;
            _currentColumn = -1; // reset column for new row

            // Return -1 when all rows are processed.
            return _currentRow <= TotalRows ? _currentRow : -1;
        }

        // No special row handling needed; method required by the interface.
        public void StartRow(Row row)
        {
            // Example: set a custom row height for the header row.
            if (_currentRow == 0)
                row.Height = 25;
        }

        // Return the next column index within the current row.
        public int NextCell()
        {
            _currentColumn++;
            return _currentColumn < TotalColumns ? _currentColumn : -1;
        }

        // Populate each cell with appropriate data.
        public void StartCell(Cell cell)
        {
            if (_currentRow == 0)
            {
                // Header row values.
                switch (cell.Column)
                {
                    case 0: cell.PutValue("ID"); break;
                    case 1: cell.PutValue("Product"); break;
                    case 2: cell.PutValue("Category"); break;
                    case 3: cell.PutValue("Price"); break;
                    case 4: cell.PutValue("Quantity"); break;
                }
            }
            else
            {
                // Sample data rows.
                switch (cell.Column)
                {
                    case 0: cell.PutValue(_currentRow); break; // ID
                    case 1: cell.PutValue($"Product {_currentRow}"); break;
                    case 2: cell.PutValue($"Category {_currentRow % 5}"); break;
                    case 3: cell.PutValue(10.0 + _currentRow * 0.5); break; // Price
                    case 4: cell.PutValue((_currentRow % 10) + 1); break; // Quantity
                }
            }
        }

        // Strings are not gathered into a global pool for this example.
        public bool IsGatherString()
        {
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and obtain the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define a merged header that spans the first three columns (A1:C1).
            // This merged area will be preserved because it is set before saving.
            cells.Merge(0, 0, 1, 3);
            // Optionally set a value for the merged header cell.
            cells[0, 0].PutValue("Sales Report");

            // Define another merged region for a sub‑header spanning columns D to E (D2:E2).
            cells.Merge(1, 3, 1, 2);
            cells[1, 3].PutValue("Summary");

            // Prepare save options with the custom LightCellsDataProvider.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new LargeTableDataProvider()
            };

            // Save the workbook. The provider streams the large table data,
            // while the previously defined merged cells are retained.
            workbook.Save("LargeTableWithMergedCells.xlsx", saveOptions);

            Console.WriteLine("Workbook saved successfully with merged cells and streamed data.");
        }
    }
}