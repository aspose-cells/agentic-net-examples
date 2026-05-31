using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace LightCellsCsvExport
{
    // Custom LightCellsDataHandler that writes each processed row to a CSV file.
    class CsvExportHandler : LightCellsDataHandler
    {
        private readonly StreamWriter _writer;
        private readonly StringBuilder _rowBuilder = new StringBuilder();
        private int _currentRowIndex = -1;

        public CsvExportHandler(string csvFilePath)
        {
            // Initialize the writer (overwrite if exists)
            _writer = new StreamWriter(csvFilePath, false, Encoding.UTF8);
        }

        // Called when a worksheet starts processing.
        public bool StartSheet(Worksheet sheet)
        {
            // Process all sheets.
            return true;
        }

        // Called before processing a row. Return true to process the row.
        public bool StartRow(int rowIndex)
        {
            // If we have data from the previous row, write it out before starting a new one.
            if (_rowBuilder.Length > 0)
            {
                _writer.WriteLine(_rowBuilder.ToString().TrimEnd(',')); // Remove trailing comma
                _rowBuilder.Clear();
            }

            _currentRowIndex = rowIndex;
            return true; // Continue processing this row
        }

        // Called after a row is started. Not needed for CSV export, just return true.
        public bool ProcessRow(Row row)
        {
            return true;
        }

        // Called before each cell in the current row. Return true to process the cell.
        public bool StartCell(int columnIndex)
        {
            // Always process every cell.
            return true;
        }

        // Called for each cell that needs to be processed.
        public bool ProcessCell(Cell cell)
        {
            // Retrieve the cell value as string, handling nulls.
            string value = cell.Value?.ToString() ?? string.Empty;

            // Escape double quotes by doubling them and wrap the value in quotes if it contains commas or quotes.
            if (value.Contains(",") || value.Contains("\""))
            {
                value = $"\"{value.Replace("\"", "\"\"")}\"";
            }

            // Append the value and a comma separator.
            _rowBuilder.Append(value).Append(',');

            // Return false to avoid keeping the cell in memory after processing.
            return false;
        }

        // Dispose the writer and flush any remaining data.
        public void Close()
        {
            if (_rowBuilder.Length > 0)
            {
                _writer.WriteLine(_rowBuilder.ToString().TrimEnd(','));
                _rowBuilder.Clear();
            }

            _writer.Flush();
            _writer.Dispose();
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook (can be any supported format, e.g., .xlsx, .csv, etc.)
            string sourceFile = "LargeData.xlsx";

            // Path where the extracted CSV will be saved.
            string outputCsv = "ExtractedData.csv";

            // Create the custom handler.
            var csvHandler = new CsvExportHandler(outputCsv);

            // Set load options to use LightCells mode with our handler.
            LoadOptions loadOptions = new LoadOptions
            {
                LightCellsDataHandler = csvHandler
            };

            // Load the workbook using LightCells. The handler will be invoked during loading.
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // After loading is complete, close the handler to flush remaining data.
            csvHandler.Close();

            Console.WriteLine($"CSV export completed: {outputCsv}");
        }
    }
}