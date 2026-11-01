// Title: How to validate column data types of an Excel ListObject table against .NET types using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells validate Excel table column data types before importing to database | Check if Excel ListObject column values match expected .NET types with Aspose.Cells | Validate data types of rows in an Excel table using Aspose.Cells in C# | Aspose.Cells column type verification for int, double, DateTime, and string columns | Report mismatched Excel column values when processing with Aspose.Cells
// Tags: Aspose.Cells column type validation | Excel ListObject data type verification C# | C# .NET type checking for Excel tables | pre‑import Excel data type validation | detect column type mismatches Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsValidation
{
    // The example loads an Excel workbook with Aspose.Cells, accesses a specified ListObject table, maps expected .NET types to column names, iterates through each data row, validates each cell against its expected type (int, double, DateTime, string), collects any mismatches, and outputs either a success message or detailed validation errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be validated
            string excelPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: File '{excelPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Name of the table (ListObject) inside the worksheet
                string tableName = "Table1";

                // Retrieve the table; if not found, abort
                ListObject table = worksheet.ListObjects[tableName];
                if (table == null)
                {
                    Console.WriteLine($"Table '{tableName}' not found in the worksheet.");
                    return;
                }

                // Define the expected .NET types for each column (key = column name, value = expected Type)
                var expectedColumnTypes = new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase)
                {
                    { "Id", typeof(int) },
                    { "Name", typeof(string) },
                    { "BirthDate", typeof(DateTime) },
                    { "Salary", typeof(double) }
                };

                // Build a map of column index -> expected Type based on the table header row
                var columnIndexToExpectedType = new Dictionary<int, Type>();
                int headerRow = table.DataRange.FirstRow - 1; // Header is one row above the data range
                for (int col = table.DataRange.FirstColumn; col < table.DataRange.FirstColumn + table.DataRange.ColumnCount; col++)
                {
                    string columnName = worksheet.Cells[headerRow, col].StringValue?.Trim();
                    if (string.IsNullOrEmpty(columnName))
                        continue;

                    if (expectedColumnTypes.TryGetValue(columnName, out Type expectedType))
                    {
                        columnIndexToExpectedType[col] = expectedType;
                    }
                    else
                    {
                        // Column not defined in expected types; you may choose to ignore or treat as error
                        Console.WriteLine($"Warning: Column '{columnName}' has no expected type defined.");
                    }
                }

                // List to collect validation errors
                var validationErrors = new List<string>();

                // Iterate through each data row in the table
                int firstDataRow = table.DataRange.FirstRow;
                int lastDataRow = table.DataRange.FirstRow + table.DataRange.RowCount - 1;

                for (int row = firstDataRow; row <= lastDataRow; row++)
                {
                    foreach (var kvp in columnIndexToExpectedType)
                    {
                        int col = kvp.Key;
                        Type expectedType = kvp.Value;

                        Cell cell = worksheet.Cells[row, col];
                        object cellValue = cell.Value;

                        // Treat empty cells as null (you may adjust this rule)
                        if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
                            continue;

                        bool isValid = false;

                        if (expectedType == typeof(int))
                        {
                            if (cellValue is int) isValid = true;
                            else if (cellValue is double d && d % 1 == 0) isValid = true;
                            else if (int.TryParse(cellValue.ToString(), out _)) isValid = true;
                        }
                        else if (expectedType == typeof(double))
                        {
                            if (cellValue is double) isValid = true;
                            else if (double.TryParse(cellValue.ToString(), out _)) isValid = true;
                        }
                        else if (expectedType == typeof(DateTime))
                        {
                            if (cellValue is DateTime) isValid = true;
                            else if (DateTime.TryParse(cellValue.ToString(), out _)) isValid = true;
                        }
                        else if (expectedType == typeof(string))
                        {
                            // Anything can be represented as string
                            isValid = true;
                        }
                        else
                        {
                            // Fallback to direct type check
                            isValid = expectedType.IsInstanceOfType(cellValue);
                        }

                        if (!isValid)
                        {
                            string columnName = worksheet.Cells[headerRow, col].StringValue;
                            validationErrors.Add($"Row {row + 1}, Column '{columnName}': Value '{cellValue}' is not of expected type {expectedType.Name}.");
                        }
                    }
                }

                // Output validation results
                if (validationErrors.Count == 0)
                {
                    Console.WriteLine("All data rows match the expected .NET types.");
                }
                else
                {
                    Console.WriteLine("Data type validation errors found:");
                    foreach (string error in validationErrors)
                    {
                        Console.WriteLine(error);
                    }
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors to prevent the application from crashing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
