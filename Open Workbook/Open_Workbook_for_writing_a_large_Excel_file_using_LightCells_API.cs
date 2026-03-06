using System;
using Aspose.Cells;

namespace LightCellsDemo
{
    // Custom data provider that streams a large amount of data to the workbook during save.
    public class LargeDataProvider : LightCellsDataProvider
    {
        private const int TotalRows = 100_000;   // Number of rows to generate
        private const int TotalCols = 10;        // Number of columns per row

        private int _currentRow = -1;
        private int _currentCol = -1;
        private bool _processSheet = false;

        // Called at the beginning of each worksheet save.
        public bool StartSheet(int sheetIndex)
        {
            // Process only the first worksheet (index 0).
            _processSheet = sheetIndex == 0;
            return _processSheet;
        }

        // Returns the next row index to be saved, or -1 when no more rows.
        public int NextRow()
        {
            if (!_processSheet) return -1;

            _currentRow++;
            _currentCol = -1; // Reset column index for the new row.

            return _currentRow < TotalRows ? _currentRow : -1;
        }

        // Called after NextRow() returns a valid index; can set row properties here.
        public void StartRow(Row row)
        {
            // Example: set a uniform row height.
            row.Height = 15;
        }

        // Returns the next column index within the current row, or -1 when the row is finished.
        public int NextCell()
        {
            _currentCol++;
            return _currentCol < TotalCols ? _currentCol : -1;
        }

        // Called after NextCell() returns a valid index; set the cell value.
        public void StartCell(Cell cell)
        {
            // Populate the cell with a simple string indicating its position.
            cell.PutValue($"R{_currentRow + 1}C{_currentCol + 1}");
        }

        // Determines whether string values should be gathered into a global string pool.
        public bool IsGatherString()
        {
            // For this example we do not need a global string pool.
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            // Initialize a new workbook instance.
            Workbook workbook = new Workbook();

            // Optionally clear existing worksheets and add a single empty sheet.
            workbook.Worksheets.Clear();
            workbook.Worksheets.Add();

            // Configure save options to use LightCells mode with the custom data provider.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new LargeDataProvider()
            };

            // Save the workbook; the data provider streams the large dataset during this call.
            workbook.Save("LargeFile.xlsx", saveOptions);

            Console.WriteLine("Large Excel file generated successfully using LightCells API.");
        }
    }
}