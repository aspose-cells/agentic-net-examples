using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsColumnTypeValidation
{
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains the source table
            // (Replace \"input.xlsx\" with the actual path to your Excel file)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];

            // Define the expected .NET types for each column (by column header name)
            // Example: "ID" should be Int32, "Name" should be String, "BirthDate" should be DateTime
            var expectedColumnTypes = new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase)
            {
                { "ID", typeof(int) },
                { "Name", typeof(string) },
                { "BirthDate", typeof(DateTime) },
                { "Salary", typeof(decimal) }
            };

            // Export the worksheet data to a DataTable.
            // CheckMixedValueType = true forces Aspose.Cells to examine all rows
            // and set the DataColumn type accordingly (or string if mixed).
            ExportTableOptions exportOptions = new ExportTableOptions
            {
                ExportColumnName = true,      // First row contains column names
                CheckMixedValueType = true    // Examine all rows for type detection
            };

            // Export the data range (adjust row/column count as needed)
            DataTable dataTable = sheet.Cells.ExportDataTable(0, 0, sheet.Cells.MaxDataRow + 1,
                                                             sheet.Cells.MaxDataColumn + 1, exportOptions);

            // Validate each column's detected type against the expected type
            foreach (DataColumn column in dataTable.Columns)
            {
                // Skip columns that are not part of the validation dictionary
                if (!expectedColumnTypes.TryGetValue(column.ColumnName, out Type expectedType))
                {
                    Console.WriteLine($"Column \"{column.ColumnName}\" is not defined in expected types – skipping.");
                    continue;
                }

                // Compare the detected DataColumn.DataType with the expected .NET type
                if (column.DataType != expectedType)
                {
                    Console.WriteLine($"Type mismatch in column \"{column.ColumnName}\": " +
                                      $"expected {expectedType.Name}, detected {column.DataType.Name}.");
                    // Here you could decide to abort, convert, or handle the mismatch as needed.
                    // For demonstration we abort the import process.
                    Console.WriteLine("Aborting import due to type mismatch.");
                    return;
                }
                else
                {
                    Console.WriteLine($"Column \"{column.ColumnName}\" type validated as {column.DataType.Name}.");
                }
            }

            // All columns validated – proceed with importing the data into the database.
            // Example: using ImportData to write the DataTable back to another worksheet
            // (replace with actual DB import logic as required).
            Worksheet targetSheet = workbook.Worksheets.Add("ValidatedData");
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,
                InsertRows = true,
                ConvertNumericData = true
            };
            targetSheet.Cells.ImportData(dataTable, 0, 0, importOptions);

            // Save the workbook with the validated data (optional)
            workbook.Save("validated_output.xlsx");
            Console.WriteLine("Data validated and exported successfully.");
        }
    }
}