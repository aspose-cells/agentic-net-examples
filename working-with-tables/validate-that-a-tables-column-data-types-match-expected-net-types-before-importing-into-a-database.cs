// Title: Validate Excel table column data types against .NET types using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells validate Excel table column data types before database import | how to check that Excel ListObject columns match expected .NET types with Aspose.Cells | Aspose.Cells C# verify integer columns contain whole numbers in an Excel table | detect type mismatches in Excel table columns using Aspose.Cells for .NET | validate Excel worksheet table data types against a dictionary of .NET types in C# | Aspose.Cells column data type verification for database loading
// Tags: Aspose.Cells column type validation | Excel ListObject data type checking | C# .NET type verification for Excel columns | integer detection in Aspose.Cells numeric cells | type mismatch reporting with Aspose.Cells | Excel to database column type verification

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject and ListColumn types
using AsposeRange = Aspose.Cells.Range; // Resolve ambiguity with System.Range

// The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet's first ListObject (table), and validates each column's cell values against a predefined dictionary of expected .NET types. Empty cells are treated as valid, numeric values returned as double are examined for integer expectations, and any type mismatches are collected and reported, providing a clear validation summary before importing data into a database.
class Program
{
    // Expected .NET types for each column (column name -> expected Type)
    static readonly Dictionary<string, Type> ExpectedColumnTypes = new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase)
    {
        { "Id", typeof(int) },
        { "Name", typeof(string) },
        { "Price", typeof(double) },
        { "CreatedDate", typeof(DateTime) }
    };

    static void Main()
    {
        const string inputPath = "Input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the first worksheet contains the table
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found in the worksheet.");
                return;
            }

            // Get the first table
            ListObject table = sheet.ListObjects[0];

            // Data range of the table (excluding header row)
            AsposeRange dataRange = table.DataRange;
            int startRow = dataRange.FirstRow;          // zero‑based index of first data row
            int startColumn = dataRange.FirstColumn;    // zero‑based index of first data column
            int rowCount = dataRange.RowCount;
            int columnCount = dataRange.ColumnCount;

            // Collect validation errors
            List<string> validationErrors = new List<string>();

            // Iterate through each column defined in the table
            for (int colIndex = 0; colIndex < columnCount; colIndex++)
            {
                ListColumn column = table.ListColumns[colIndex];
                string columnName = column.Name;

                // Skip columns that are not defined in the expected type dictionary
                if (!ExpectedColumnTypes.TryGetValue(columnName, out Type expectedType))
                    continue;

                bool columnValid = true;

                // Scan each data row for the current column
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                {
                    // Retrieve the cell value for the current row and column
                    Cell cell = sheet.Cells[startRow + rowIndex, startColumn + colIndex];
                    object value = cell.Value;

                    // Treat empty cells as valid (they can be null or empty string)
                    if (value == null || string.IsNullOrEmpty(value.ToString()))
                        continue;

                    Type actualType = value.GetType();

                    // Aspose.Cells returns double for all numeric values; handle integer expectations specially
                    if (expectedType == typeof(int) && actualType == typeof(double))
                    {
                        double d = (double)value;
                        if (d % 1 != 0) // non‑integer value found
                        {
                            columnValid = false;
                            validationErrors.Add($"Column '{columnName}' expects integer but found non‑integer value '{value}' at row {startRow + rowIndex + 2}.");
                            break;
                        }
                    }
                    else if (!expectedType.IsAssignableFrom(actualType))
                    {
                        columnValid = false;
                        validationErrors.Add($"Column '{columnName}' expects {expectedType.Name} but found {actualType.Name} value '{value}' at row {startRow + rowIndex + 2}.");
                        break;
                    }
                }

                if (columnValid)
                {
                    Console.WriteLine($"Column '{columnName}' passed validation.");
                }
            }

            // Output validation summary
            if (validationErrors.Count == 0)
            {
                Console.WriteLine("All columns match the expected .NET types.");
            }
            else
            {
                Console.WriteLine("Validation errors found:");
                foreach (string err in validationErrors)
                {
                    Console.WriteLine("- " + err);
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors (e.g., file format issues, Aspose.Cells exceptions)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
