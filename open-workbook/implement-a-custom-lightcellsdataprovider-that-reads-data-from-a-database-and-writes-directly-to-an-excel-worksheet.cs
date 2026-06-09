using System;
using System.IO;
using Aspose.Cells;

namespace LightCellsDbExport
{
    // Simple LightCellsDataProvider that streams static data into the workbook.
    public class DbLightCellsDataProvider : LightCellsDataProvider, IDisposable
    {
        // Sample data to export (you can replace this with real DB logic).
        private readonly string[,] _data = new string[,]
        {
            { "Id", "Name", "Price", "Qty" },
            { "1", "Apple", "0.5", "100" },
            { "2", "Banana", "0.3", "150" },
            { "3", "Orange", "0.7", "80" }
        };

        private int _rowCount;
        private int _colCount;
        private int _currentRow = -1;      // Zero‑based row index being processed
        private int _currentColumn = -1;   // Zero‑based column index being processed

        public DbLightCellsDataProvider()
        {
            _rowCount = _data.GetLength(0);
            _colCount = _data.GetLength(1);
        }

        // Called once for each worksheet. We only handle the first sheet (index 0).
        public bool StartSheet(int sheetIndex)
        {
            return sheetIndex == 0;
        }

        // Returns the next row index to be saved, or -1 when no more rows exist.
        public int NextRow()
        {
            try
            {
                if (_currentRow + 1 < _rowCount)
                {
                    _currentRow++;
                    _currentColumn = -1; // reset column for the new row
                    return _currentRow;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error reading next row: {ex.Message}");
            }
            return -1; // No more rows or error
        }

        // Allows setting row‑level properties (height, style, etc.). Not needed for this demo.
        public void StartRow(Row row)
        {
            // Example: set a fixed height for all rows
            // row.Height = 15;
        }

        // Returns the next column index for the current row, or -1 when the row is finished.
        public int NextCell()
        {
            if (_currentColumn < _colCount - 1)
            {
                _currentColumn++;
                return _currentColumn;
            }
            return -1; // End of columns for this row
        }

        // Writes the value from the data array into the cell.
        public void StartCell(Cell cell)
        {
            try
            {
                object value = _data[_currentRow, _currentColumn];
                cell.PutValue(value);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error writing cell value: {ex.Message}");
            }
        }

        // Indicates whether string values should be gathered into the global string pool.
        public bool IsGatherString()
        {
            return true;
        }

        // Clean up resources (none needed for static data).
        public void Dispose()
        {
            // No unmanaged resources to release.
        }
    }

    class Program
    {
        static void Main()
        {
            // Output file path
            string outputPath = "DatabaseExport.xlsx";

            try
            {
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Create the custom provider.
                using var provider = new DbLightCellsDataProvider();

                // Configure save options to use LightCells mode with our provider.
                var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
                {
                    LightCellsDataProvider = provider
                };

                // Create an empty workbook (no data model needed because data comes from the provider).
                var workbook = new Workbook();

                // Save the workbook; data is streamed from the provider directly into the file.
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine("Workbook saved successfully using LightCellsDataProvider.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}