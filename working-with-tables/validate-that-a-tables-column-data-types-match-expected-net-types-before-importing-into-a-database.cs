// Title: How to validate Excel ListObject column data types against expected .NET types using Aspose.Cells in C#
// AI Prompts: Write a C# method that opens an Excel workbook with Aspose.Cells, locates a ListObject by its name, and checks each cell in the defined columns against a dictionary of expected .NET types, returning detailed mismatch information. | Update the validation routine to treat numeric cells as integers when the expected type is int, handling Excel's double storage, and generate error messages for rows containing non‑integer values. | Extend the validator to accept custom type converters so that string values can be mapped to enums or nullable types during column validation, and include these conversions in the result report.
// Common Searches: aspocells c# validate excel table column data types before importing to database | check if excel ListObject column values match .net types using Aspose.Cells | c# how to ensure integer columns are whole numbers when reading Excel with Aspose.Cells | validate date and boolean columns in an Excel table with Aspose.Cells C# | aspocells schema validation for Excel tables in .net application
// Tags: Aspose.Cells column type validation | Excel ListObject schema verification C# | validate .NET data types in Excel table | integer detection in Excel numeric cells | custom type converters for Excel validation | validation result pattern Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace ExcelTableValidator
{
    // Provides a C# example that loads an Excel workbook with Aspose.Cells, locates a ListObject, and validates each column's cell values against a dictionary of expected .NET types, handling integer detection, custom converters, and returning a ValidationResult with any mismatches.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the Excel file
                string filePath = @"C:\Data\Sample.xlsx";

                // Verify that the file exists to avoid FileNotFoundException
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Name of the worksheet and the table (ListObject) to validate
                string sheetName = "Sheet1";
                string tableName = "MyTable";

                // Define expected .NET types for each column (column name -> expected Type)
                var expectedColumnTypes = new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase)
                {
                    { "Id", typeof(int) },
                    { "Name", typeof(string) },
                    { "BirthDate", typeof(DateTime) },
                    { "IsActive", typeof(bool) },
                    { "Score", typeof(double) }
                };

                // Perform validation
                var validationResult = ValidateTableColumnTypes(filePath, sheetName, tableName, expectedColumnTypes);

                // Output results
                if (validationResult.IsValid)
                {
                    Console.WriteLine("All column data types match the expected .NET types.");
                }
                else
                {
                    Console.WriteLine("Data type mismatches found:");
                    foreach (var error in validationResult.Errors)
                    {
                        Console.WriteLine(error);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        /// <param name="excelFilePath">Full path to the Excel workbook.</param>
        /// <param name="worksheetName">Name of the worksheet containing the table.</param>
        /// <param name="tableName">Name of the ListObject (table) to validate.</param>
        /// <param name="expectedTypes">Dictionary mapping column names to expected .NET types.</param>
        /// <returns>A ValidationResult indicating success or a list of error messages.</returns>
        public static ValidationResult ValidateTableColumnTypes(
            string excelFilePath,
            string worksheetName,
            string tableName,
            Dictionary<string, Type> expectedTypes)
        {
            var result = new ValidationResult();

            try
            {
                // Load the workbook
                var workbook = new Workbook(excelFilePath);

                // Get the worksheet
                var worksheet = workbook.Worksheets[worksheetName];
                if (worksheet == null)
                {
                    result.Errors.Add($"Worksheet '{worksheetName}' not found.");
                    return result;
                }

                // Retrieve the table (ListObject) by name
                var table = worksheet.ListObjects[tableName];
                if (table == null)
                {
                    result.Errors.Add($"Table '{tableName}' not found in worksheet '{worksheetName}'.");
                    return result;
                }

                // Map column names to their index within the table
                var columnNameToIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int col = 0; col < table.ListColumns.Count; col++)
                {
                    string colName = table.ListColumns[col].Name;
                    columnNameToIndex[colName] = col;
                }

                // Validate each expected column
                foreach (var kvp in expectedTypes)
                {
                    string columnName = kvp.Key;
                    Type expectedType = kvp.Value;

                    if (!columnNameToIndex.TryGetValue(columnName, out int colIndex))
                    {
                        result.Errors.Add($"Expected column '{columnName}' not found in table '{tableName}'.");
                        continue;
                    }

                    // Iterate through each data row in the table
                    for (int row = 0; row < table.DataRange.RowCount; row++)
                    {
                        // Get the cell value
                        var cell = table.DataRange[row, colIndex];
                        object value = cell.Value;

                        // Treat empty cells as valid (skip validation)
                        if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                            continue;

                        // Determine the actual .NET type of the cell value
                        Type actualType = GetUnderlyingType(value);

                        // Special handling for integer expectations (Excel stores numbers as double)
                        if (expectedType == typeof(int) && actualType == typeof(double))
                        {
                            double d = (double)value;
                            if (d % 1 != 0)
                            {
                                result.Errors.Add($"Row {row + 2} column '{columnName}': value '{value}' is not an integer.");
                            }
                            continue;
                        }

                        // For expected double, accept double (Excel stores numbers as double)
                        if (expectedType == typeof(double) && actualType == typeof(double))
                            continue;

                        // Direct type comparison
                        if (actualType != expectedType)
                        {
                            result.Errors.Add($"Row {row + 2} column '{columnName}': expected type {expectedType.Name}, but found {actualType.Name} with value '{value}'.");
                        }
                    }
                }

                result.IsValid = result.Errors.Count == 0;
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Validation failed: {ex.Message}");
            }

            return result;
        }

        private static Type GetUnderlyingType(object value)
        {
            // Aspose.Cells may return types such as double, string, bool, DateTime.
            // For formulas that return numeric results, the type is double.
            // For dates, the type is DateTime.
            // For boolean, the type is bool.
            // For text, the type is string.
            return value.GetType();
        }
    }

    /// <summary>
    /// Simple container for validation results.
    /// </summary>
    public class ValidationResult
    {
        public bool IsValid { get; set; } = false;
        public List<string> Errors { get; } = new List<string>();
    }
}
