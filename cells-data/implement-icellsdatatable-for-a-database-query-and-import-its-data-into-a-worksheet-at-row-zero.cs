// Title: Create a custom ICellsDataTable from IDataReader and import DB results into an Aspose.Cells worksheet at row 0
// Description: Demonstrates a C# implementation of ICellsDataTable that reads all rows from an IDataReader into memory, exposes column names and indexers, and uses Worksheet.Cells.ImportData with ImportTableOptions (IsFieldNameShown = true) to place the data at the first row and column of a new workbook, then saves the file as an Excel document.
// Keywords: Aspose.Cells | ICellsDataTable | IDataReader | ImportData | ImportTableOptions | C# Excel export | database query to Excel | custom data table | .NET workbook generation | row 0 import
// Common Searches: how to implement ICellsDataTable from IDataReader | import database query into Aspose.Cells worksheet | show column headers when importing data with Aspose.Cells | Aspose.Cells import data starting at first row | C# export IDataReader to Excel using Aspose
// Developer Intent: Build a reusable ICellsDataTable wrapper for IDataReader and load its content into an Excel sheet beginning at the top‑left cell.
// Use Cases: Export results of a SQL SELECT command directly to an Excel file without intermediate CSV files. | Reuse the same IDataReader‑based table to populate multiple worksheets or different start positions in a workbook. | Create a lightweight data‑export utility for .NET applications that need to generate reports from in‑memory data sources.
// AI Prompts: Generate a C# ICellsDataTable implementation that streams rows from an IDataReader instead of pre‑loading them. | Show how to import a custom ICellsDataTable into cell range B2:C10 and apply a built‑in table style with Aspose.Cells. | Explain handling of nullable database columns in the ICellsDataTable indexer for robust Excel export.

using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsICellsDataTableDemo
{
    // Custom implementation of ICellsDataTable that wraps an IDataReader.
    // It loads all data into memory for random access during import.
    // Demonstrates a C# implementation of ICellsDataTable that reads all rows from an IDataReader into memory, exposes column names and indexers, and uses Worksheet.Cells.ImportData with ImportTableOptions (IsFieldNameShown = true) to place the data at the first row and column of a new workbook, then saves the file as an Excel document.
    public class DataReaderCellsDataTable : ICellsDataTable
    {
        private readonly List<object[]> _rows = new List<object[]>();
        private readonly string[] _columns;
        private int _currentRow = -1;

        // Constructor reads the IDataReader, captures column names and all rows.
        public DataReaderCellsDataTable(IDataReader reader)
        {
            // Capture column names.
            var columnList = new List<string>();
            for (int i = 0; i < reader.FieldCount; i++)
            {
                columnList.Add(reader.GetName(i));
            }
            _columns = columnList.ToArray();

            // Read all rows.
            while (reader.Read())
            {
                var row = new object[reader.FieldCount];
                reader.GetValues(row);
                _rows.Add(row);
            }
        }

        // Column names.
        public string[] Columns => _columns;

        // Number of data rows (excluding header).
        public int Count => _rows.Count;

        // Indexer by column index for the current row.
        public object this[int column] => _rows[_currentRow][column];

        // Indexer by column name for the current row.
        public object this[string columnName]
        {
            get
            {
                int idx = Array.IndexOf(_columns, columnName);
                if (idx < 0) throw new ArgumentException($"Column '{columnName}' does not exist.");
                return this[idx];
            }
        }

        // Reset cursor to before the first row.
        public void BeforeFirst()
        {
            _currentRow = -1;
        }

        // Move to the next row; return false if no more rows.
        public bool Next()
        {
            _currentRow++;
            return _currentRow < _rows.Count;
        }
    }

    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a sample DataTable to simulate a DB query.
            // -------------------------------------------------
            DataTable sampleTable = new DataTable("Employees");
            sampleTable.Columns.Add("Id", typeof(int));
            sampleTable.Columns.Add("Name", typeof(string));
            sampleTable.Columns.Add("Department", typeof(string));

            sampleTable.Rows.Add(1, "John Doe", "Engineering");
            sampleTable.Rows.Add(2, "Jane Smith", "Marketing");
            sampleTable.Rows.Add(3, "Mike Johnson", "Sales");

            // -------------------------------------------------
            // 2. Obtain an IDataReader from the DataTable.
            // -------------------------------------------------
            using (IDataReader reader = sampleTable.CreateDataReader())
            {
                // -------------------------------------------------
                // 3. Wrap the IDataReader with our custom ICellsDataTable.
                // -------------------------------------------------
                ICellsDataTable cellsDataTable = new DataReaderCellsDataTable(reader);

                // -------------------------------------------------
                // 4. Create a workbook and import the data at row 0, column 0.
                // -------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    // Show column headers as the first row.
                    IsFieldNameShown = true
                };

                sheet.Cells.ImportData(cellsDataTable, 0, 0, importOptions);

                // -------------------------------------------------
                // 5. Save the workbook.
                // -------------------------------------------------
                workbook.Save("DatabaseQueryImport.xlsx");
            }
        }
    }
}
