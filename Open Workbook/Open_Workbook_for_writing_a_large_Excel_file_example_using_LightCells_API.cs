using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace LightCellsExample
{
    // Custom provider that streams data to the workbook during save
    public class LargeDataProvider : LightCellsDataProvider
    {
        private const int TotalRows = 10000;      // example large row count
        private const int TotalColumns = 50;      // example column count
        private int _currentRow = -1;
        private int _currentColumn = -1;
        private bool _processSheet = false;

        // Called once per worksheet; we process only the first sheet (index 0)
        public bool StartSheet(int sheetIndex)
        {
            _processSheet = sheetIndex == 0;
            return _processSheet;
        }

        // Return the next row index to be saved, or -1 when done
        public int NextRow()
        {
            if (!_processSheet) return -1;

            _currentRow++;
            _currentColumn = -1; // reset column for new row

            return _currentRow < TotalRows ? _currentRow : -1;
        }

        // Called for each row that will be saved; can set row properties here
        public void StartRow(Row row)
        {
            // Example: set a fixed height for all rows
            row.Height = 15;
        }

        // Return the next column index for the current row, or -1 when done
        public int NextCell()
        {
            _currentColumn++;
            return _currentColumn < TotalColumns ? _currentColumn : -1;
        }

        // Called for each cell; set the cell value
        public void StartCell(Cell cell)
        {
            // Simple data: "R{row}C{col}"
            cell.PutValue($"R{_currentRow}C{_currentColumn}");
        }

        // Determines whether string values should be gathered into a global pool
        public bool IsGatherString()
        {
            // For this example we do not need a global string pool
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Configure save options to use LightCells mode with our custom provider
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new LargeDataProvider()
            };

            // Save the workbook; the data will be streamed via the provider
            workbook.Save("LargeFile_LightCells.xlsx", saveOptions);

            Console.WriteLine("Large workbook saved using LightCells API.");
        }
    }
}