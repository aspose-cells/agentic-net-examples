// Title: Implement ICellsDataTable for a DataTable and import its rows into an Excel worksheet at row 0 using Aspose.Cells for .NET
// Common Searches: C# implement ICellsDataTable to wrap ADO.NET DataTable for Aspose.Cells ImportData | Import DataTable rows into Excel starting at first row with Aspose.Cells Cells.ImportData | Read‑only ICellsDataTable example for exporting database query results to XLSX | Aspose.Cells ImportData options to include column headers from ICellsDataTable | Wrap System.Data.DataTable as ICellsDataTable for bulk Excel export in .NET
// Tags: ICellsDataTable implementation for DataTable | Aspose.Cells ImportData from ICellsDataTable | export ADO.NET DataTable to XLSX | read‑only ICellsDataTable wrapper | import data starting at row zero Aspose.Cells

using System;
using System.Data;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsICellsDataTableDemo
{
    // Implementation of ICellsDataTable that wraps a System.Data.DataTable.
    // This allows us to import data retrieved from a database query directly into a worksheet.
    // The sample defines DatabaseCellsDataTable, a read‑only ICellsDataTable that wraps a System.Data.DataTable. After populating a DataTable with sample query results, the wrapper is passed to Cells.ImportData with ImportTableOptions (IsFieldNameShown = true). The data, including column headers, is written to the first worksheet starting at cell A1 (row 0, column 0) and saved as an XLSX file.
    public class DatabaseCellsDataTable : ICellsDataTable
    {
        private readonly DataTable _dataTable;
        private int _currentRow = -1; // Position for enumeration

        public DatabaseCellsDataTable(DataTable dataTable)
        {
            _dataTable = dataTable ?? throw new ArgumentNullException(nameof(dataTable));
        }

        // ICellsDataTable members

        // Number of rows.
        public int Count => _dataTable.Rows.Count;

        // Column names required by ICellsDataTable.
        public string[] Columns => _dataTable.Columns.Cast<DataColumn>()
                                                    .Select(c => c.ColumnName)
                                                    .ToArray();

        // Reset enumeration to before the first row.
        public void BeforeFirst()
        {
            _currentRow = -1;
        }

        // Move to the next row; return false if no more rows.
        public bool Next()
        {
            _currentRow++;
            return _currentRow < _dataTable.Rows.Count;
        }

        // Get value by column index for the current row.
        public object this[int column]
        {
            get
            {
                if (_currentRow < 0 || _currentRow >= Count)
                    throw new InvalidOperationException("Enumerator is not positioned on a valid row.");
                if (column < 0 || column >= Columns.Length)
                    throw new ArgumentOutOfRangeException(nameof(column));
                return _dataTable.Rows[_currentRow][column];
            }
        }

        // Get value by column name for the current row.
        public object this[string columnName]
        {
            get
            {
                if (_currentRow < 0 || _currentRow >= Count)
                    throw new InvalidOperationException("Enumerator is not positioned on a valid row.");
                if (string.IsNullOrEmpty(columnName))
                    throw new ArgumentNullException(nameof(columnName));
                return _dataTable.Rows[_currentRow][columnName];
            }
        }

        // Return the name of a column by index.
        public string GetColumnName(int column)
        {
            if (column < 0 || column >= Columns.Length)
                throw new ArgumentOutOfRangeException(nameof(column));
            return _dataTable.Columns[column].ColumnName;
        }

        // Return the .NET type of a column by index.
        public Type GetColumnType(int column)
        {
            if (column < 0 || column >= Columns.Length)
                throw new ArgumentOutOfRangeException(nameof(column));
            return _dataTable.Columns[column].DataType;
        }

        // The table is read‑only.
        public bool IsReadOnly => true;

        public void SetValue(int row, int column, object value)
        {
            throw new NotSupportedException("DatabaseCellsDataTable is read‑only.");
        }

        public void AddColumn(string columnName, Type columnType)
        {
            throw new NotSupportedException("Columns cannot be added to a read‑only DatabaseCellsDataTable.");
        }

        public void AddRow(object[] values)
        {
            throw new NotSupportedException("Rows cannot be added to a read‑only DatabaseCellsDataTable.");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a DataTable and populate it with sample data (simulating a DB query).
                DataTable dt = new DataTable();
                dt.Columns.Add("Id", typeof(int));
                dt.Columns.Add("Name", typeof(string));
                dt.Columns.Add("Age", typeof(int));
                dt.Columns.Add("Department", typeof(string));

                dt.Rows.Add(1, "John Doe", 30, "Sales");
                dt.Rows.Add(2, "Jane Smith", 28, "Marketing");
                dt.Rows.Add(3, "Sam Brown", 35, "IT");

                // 2. Wrap the DataTable with our ICellsDataTable implementation.
                ICellsDataTable cellsDataTable = new DatabaseCellsDataTable(dt);

                // 3. Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 4. Import the data starting at row 0, column 0.
                ImportTableOptions options = new ImportTableOptions
                {
                    IsFieldNameShown = true // Write column headers.
                };

                // ImportData works with ICellsDataTable.
                cells.ImportData(cellsDataTable, 0, 0, options);

                // 5. Save the workbook to disk.
                string outputPath = "DatabaseImportDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.Error.WriteLine("An error occurred:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
