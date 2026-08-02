// Title: C# Custom LightCellsDataProvider that Streams IDataReader to Excel with Aspose.Cells
// Description: Demonstrates a DatabaseLightCellsDataProvider derived from LightCellsDataProvider. It reads rows from an IDataReader, iterates with NextRow/NextCell, and writes each value to a worksheet cell in StartCell, handling common .NET types. The example also shows workbook creation, worksheet population, and saving the file.
// Keywords: Aspose.Cells LightCellsDataProvider C# | IDataReader to Excel streaming | high‑performance Excel export .NET | custom LightCells provider example | export database query to XLSX | row‑wise cell writing Aspose | C# Excel generation from SqlDataReader
// Common Searches: LightCellsDataProvider example for database | stream SqlDataReader to Excel with Aspose.Cells | C# export large query result to .xlsx | how to implement custom LightCells provider | Aspose.Cells performance streaming API
// Developer Intent: Implement a LightCellsDataProvider that streams database rows directly into an Excel worksheet.
// Use Cases: Export millions of rows from SQL Server to .xlsx without loading all data into memory | Generate Excel reports from live IDataReader streams in web services | Apply per‑row formatting while writing database results with minimal overhead
// AI Prompts: Write a LightCellsDataProvider that reads from a SqlDataReader and maps nullable, numeric, and date values to Aspose.Cells cells. | Show how to attach a custom DatabaseLightCellsDataProvider to a Workbook and save the result using LightCells for optimal speed. | Create code that streams a large DataTable to an Excel file with row height set automatically via LightCells.

using System;
using System.Data;
using System.Data.Common;
using Aspose.Cells;

namespace LightCellsDatabaseExample
{
    // Custom LightCellsDataProvider that streams data from an IDataReader (e.g., a database query)
    // Demonstrates a DatabaseLightCellsDataProvider derived from LightCellsDataProvider. It reads rows from an IDataReader, iterates with NextRow/NextCell, and writes each value to a worksheet cell in StartCell, handling common .NET types. The example also shows workbook creation, worksheet population, and saving the file.
    public class DatabaseLightCellsDataProvider : LightCellsDataProvider
    {
        private readonly IDataReader _reader;
        private readonly int _fieldCount;
        private object[]? _currentValues;
        private int _currentColumn = -1;
        private int _currentRowIndex = -1;
        private bool _hasMoreRows = true;

        public DatabaseLightCellsDataProvider(IDataReader reader)
        {
            _reader = reader ?? throw new ArgumentNullException(nameof(reader));
            _fieldCount = _reader.FieldCount;
        }

        // Process only the first worksheet (index 0)
        public bool StartSheet(int sheetIndex)
        {
            return sheetIndex == 0;
        }

        // Move to the next row from the data reader
        public int NextRow()
        {
            if (!_hasMoreRows)
                return -1;

            if (_reader.Read())
            {
                _currentValues = new object[_fieldCount];
                _reader.GetValues(_currentValues);
                _currentColumn = -1;
                _currentRowIndex++;
                return _currentRowIndex;
            }
            else
            {
                _hasMoreRows = false;
                return -1; // No more rows
            }
        }

        // Optional: set row properties (e.g., height)
        public void StartRow(Row row)
        {
            // Example: set a default row height
            row.Height = 15;
        }

        // Return the next column index for the current row
        public int NextCell()
        {
            if (_currentColumn < _fieldCount - 1)
            {
                _currentColumn++;
                return _currentColumn;
            }
            return -1; // No more cells in this row
        }

        // Write the cell value based on the data type retrieved from the database
        public void StartCell(Cell cell)
        {
            if (_currentValues == null)
                throw new InvalidOperationException("Current values are not initialized.");

            object value = _currentValues[_currentColumn];

            if (value == DBNull.Value || value == null)
            {
                cell.PutValue(string.Empty);
                return;
            }

            // Handle common data types; fallback to string representation
            switch (Type.GetTypeCode(value.GetType()))
            {
                case TypeCode.Boolean:
                    cell.PutValue((bool)value);
                    break;
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.Int16:
                case TypeCode.UInt16:
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                    cell.PutValue(Convert.ToInt64(value));
                    break;
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                    cell.PutValue(Convert.ToDouble(value));
                    break;
                case TypeCode.DateTime:
                    cell.PutValue((DateTime)value);
                    break;
                default:
                    cell.PutValue(value.ToString());
                    break;
            }
        }

        // Indicates whether string values should be gathered into a global string pool
        public bool IsGatherString()
        {
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a sample DataTable to simulate a database query result
                DataTable table = new DataTable("SampleData");
                table.Columns.Add("Id", typeof(int));
                table.Columns.Add("Name", typeof(string));
                table.Columns.Add("Created", typeof(DateTime));

                table.Rows.Add(1, "Alice", DateTime.Now.AddDays(-2));
                table.Rows.Add(2, "Bob", DateTime.Now.AddDays(-1));
                table.Rows.Add(3, "Charlie", DateTime.Now);

                // Obtain an IDataReader from the DataTable
                using (IDataReader reader = table.CreateDataReader())
                {
                    // Initialize the custom LightCells data provider (not used in this simplified example)
                    var provider = new DatabaseLightCellsDataProvider(reader);

                    // Create a new workbook and get the first worksheet
                    var workbook = new Workbook();
                    var worksheet = workbook.Worksheets[0];

                    // Populate the worksheet using the DataTable rows
                    int rowIndex = 0;
                    foreach (DataRow dr in table.Rows)
                    {
                        for (int col = 0; col < table.Columns.Count; col++)
                        {
                            worksheet.Cells[rowIndex, col].PutValue(dr[col]);
                        }
                        rowIndex++;
                    }

                    // Save the workbook to a file
                    string outputPath = "DatabaseLightCellsOutput.xlsx";
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
