// Title: Import a CSV file into an Aspose.Cells worksheet using a custom ICellsDataTable implementation (C#)
// Common Searches: C# how to use Aspose.Cells ImportData with a custom ICellsDataTable to load CSV data | sample code for importing CSV into an Excel workbook using Aspose.Cells ImportTableOptions | read CSV file into Aspose.Cells worksheet with column headers using a custom data table | Aspose.Cells ImportData example with field name handling for CSV files in C# | create a workbook from CSV using ICellsDataTable implementation in Aspose.Cells
// Tags: CSV data source via ICellsDataTable | Aspose.Cells ImportData CSV example | ImportTableOptions field name flag | Workbook save Xlsx after import | reading CSV into worksheet cells Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsCsvImport
{
    // Custom implementation of ICellsDataTable to read CSV data
    // The example defines a CsvDataTable class that implements ICellsDataTable to read a CSV file, exposing rows, columns, and column‑name indexing. A Workbook is created, the first worksheet’s Cells collection is obtained, and Cells.ImportData is called with the custom table and ImportTableOptions (IsFieldNameShown = true) to import the CSV starting at cell A1. After importing, the workbook is saved as an XLSX file. If the CSV file is missing, a sample CSV is generated automatically.
    public class CsvDataTable : ICellsDataTable
    {
        private readonly List<string[]> _rows;
        private readonly Dictionary<string, int> _headerIndex;
        private readonly string[] _headers;
        private int _currentRow = -1;

        public CsvDataTable(string csvFilePath)
        {
            if (!File.Exists(csvFilePath))
                throw new FileNotFoundException($"CSV file not found: {csvFilePath}");

            _rows = new List<string[]>();
            foreach (var line in File.ReadAllLines(csvFilePath))
            {
                // Simple CSV split – does not handle quoted commas.
                var fields = line.Split(',');
                _rows.Add(fields);
            }

            // Build header index for column name lookup (first row is assumed to be header)
            _headerIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (_rows.Count > 0)
            {
                _headers = _rows[0];
                for (int i = 0; i < _headers.Length; i++)
                {
                    var header = _headers[i];
                    if (!string.IsNullOrEmpty(header) && !_headerIndex.ContainsKey(header))
                        _headerIndex[header] = i;
                }
            }
            else
            {
                _headers = Array.Empty<string>();
            }
        }

        // ICellsDataTable members

        public int RowCount => _rows.Count;

        public int ColumnCount => _rows.Count > 0 ? _rows[0].Length : 0;

        public object GetValue(int row, int column)
        {
            if (row < 0 || row >= RowCount || column < 0 || column >= ColumnCount)
                return null;
            return _rows[row][column];
        }

        public Type GetDataType(int column)
        {
            // Treat all columns as string for simplicity
            return typeof(string);
        }

        public void BeforeFirst()
        {
            _currentRow = -1;
        }

        public bool Next()
        {
            _currentRow++;
            return _currentRow < RowCount;
        }

        // Returns the column names (required by ICellsDataTable)
        public string[] Columns => _headers;

        // Total number of rows (required by ICellsDataTable)
        public int Count => RowCount;

        // Indexer by row/column
        public object this[int row, int column] => GetValue(row, column);

        // Indexer for the current row by column index
        public object this[int column]
        {
            get
            {
                if (_currentRow < 0 || _currentRow >= RowCount)
                    return null;
                return GetValue(_currentRow, column);
            }
        }

        // Indexer by column name (uses the current row set by Next())
        public object this[string columnName]
        {
            get
            {
                if (!_headerIndex.TryGetValue(columnName, out int colIdx))
                    return null;
                if (_currentRow < 0 || _currentRow >= RowCount)
                    return null;
                return GetValue(_currentRow, colIdx);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Path to the CSV file (adjust as needed)
            string csvPath = "data.csv";

            // Ensure sample CSV exists
            if (!File.Exists(csvPath))
            {
                var sampleLines = new[]
                {
                    "Name,Age,Country",
                    "John,30,USA",
                    "Anna,25,UK",
                    "Lee,28,Canada"
                };
                try
                {
                    File.WriteAllLines(csvPath, sampleLines);
                    Console.WriteLine($"Sample CSV created at '{Path.GetFullPath(csvPath)}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create sample CSV: {ex.Message}");
                    return;
                }
            }

            try
            {
                // Create a new workbook and get the first worksheet's cells collection
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Build the custom CSV data table
                CsvDataTable csvTable = new CsvDataTable(csvPath);

                // Set import options – first row contains field names
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    IsFieldNameShown = true
                };

                // Import the CSV data starting at cell A1 (row 0, column 0)
                int importedRows = cells.ImportData(csvTable, 0, 0, importOptions);
                Console.WriteLine($"Imported {importedRows} rows from CSV.");

                // Save the workbook
                string outputPath = "output.xlsx";
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during import: {ex.Message}");
            }
        }
    }
}
