// Title: Validate Excel column types against .NET types before generating a SQL script with Aspose.Cells for .NET
// Description: Shows how to build a workbook, export a range to a DataTable using ExportTableOptions with mixed‑value type checking, compare each DataColumn.DataType to a predefined .NET type map, and create a CREATE TABLE/INSERT SQL script only when all columns match, leveraging SqlScriptSaveOptions.
// Keywords: Aspose.Cells | C# | ExportDataTable | ExportTableOptions | CheckMixedValueType | DataTable column type validation | .NET type checking | SQL script generation | SqlScriptSaveOptions | Excel to SQL import | data import validation
// Common Searches: Aspose.Cells validate Excel column data type | Export worksheet to DataTable with type checking C# | Check DataColumn.DataType against .NET types | Generate SQL script from workbook after validation | How to abort SQL export when column type mismatch
// Developer Intent: Confirm that each column exported from an Excel sheet matches the expected .NET type before producing a SQL import script.
// Use Cases: Export a worksheet to a DataTable while automatically detecting column types. | Validate DataTable columns against a dictionary of expected .NET types and log mismatches. | Prevent SQL script creation if any column fails the type check. | Automatically generate CREATE TABLE and INSERT statements for a verified worksheet.
// AI Prompts: Write C# code using Aspose.Cells to export a worksheet to a DataTable with mixed‑value type checking and validate each column against a dictionary of expected .NET types. | Create a method that receives a DataTable and a Dictionary<int, Type>, returns true if all column DataTypes match, and logs any mismatches. | Show how to configure SqlScriptSaveOptions to produce a CREATE TABLE and INSERT script only after successful column type validation.

using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsValidationDemo
{
    // Shows how to build a workbook, export a range to a DataTable using ExportTableOptions with mixed‑value type checking, compare each DataColumn.DataType to a predefined .NET type map, and create a CREATE TABLE/INSERT SQL script only when all columns match, leveraging SqlScriptSaveOptions.
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a workbook and populate it with sample data
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Header row
            sheet.Cells["A1"].PutValue("Id");          // Expected int
            sheet.Cells["B1"].PutValue("Name");        // Expected string
            sheet.Cells["C1"].PutValue("BirthDate");   // Expected DateTime

            // Data rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(new DateTime(1990, 5, 23));

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(new DateTime(1985, 11, 12));

            // ------------------------------------------------------------
            // 2. Export the worksheet to a DataTable with type checking
            // ------------------------------------------------------------
            ExportTableOptions exportOptions = new ExportTableOptions
            {
                ExportColumnName = true,          // First row contains column names
                CheckMixedValueType = true        // Examine all rows to determine column types
            };

            // Export all rows and columns (3 rows including header, 3 columns)
            DataTable dataTable = sheet.Cells.ExportDataTable(0, 0, 3, 3, exportOptions);

            // ------------------------------------------------------------
            // 3. Define the expected .NET types for each column (by index)
            // ------------------------------------------------------------
            var expectedColumnTypes = new Dictionary<int, Type>
            {
                { 0, typeof(int) },        // Id column
                { 1, typeof(string) },     // Name column
                { 2, typeof(DateTime) }    // BirthDate column
            };

            // ------------------------------------------------------------
            // 4. Validate that each column's DataType matches the expectation
            // ------------------------------------------------------------
            bool validationPassed = true;
            foreach (DataColumn column in dataTable.Columns)
            {
                int colIndex = column.Ordinal;
                if (expectedColumnTypes.TryGetValue(colIndex, out Type expectedType))
                {
                    if (column.DataType != expectedType)
                    {
                        Console.WriteLine($"Column '{column.ColumnName}' (index {colIndex}) type mismatch. " +
                                          $"Expected: {expectedType.FullName}, Actual: {column.DataType.FullName}");
                        validationPassed = false;
                    }
                }
                else
                {
                    Console.WriteLine($"No expected type defined for column index {colIndex}. Skipping validation.");
                }
            }

            // ------------------------------------------------------------
            // 5. If validation succeeds, generate a SQL script for import
            // ------------------------------------------------------------
            if (validationPassed)
            {
                SqlScriptSaveOptions sqlOptions = new SqlScriptSaveOptions
                {
                    OperatorType = SqlScriptOperatorType.Insert,
                    CreateTable = true,
                    TableName = "People",
                    CheckAllDataForColumnType = true   // Ensure column types are derived from all data rows
                };

                // Save the workbook as a SQL script file
                workbook.Save("PeopleData.sql", sqlOptions);
                Console.WriteLine("Validation succeeded. SQL script 'PeopleData.sql' generated.");
            }
            else
            {
                Console.WriteLine("Validation failed. SQL script generation aborted.");
            }
        }
    }
}
