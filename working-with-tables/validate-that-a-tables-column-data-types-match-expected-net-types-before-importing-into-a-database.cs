// Title: Validate Excel Column Types Against .NET Types Before Database Import with Aspose.Cells (C#)
// Description: This example creates a workbook, defines the expected .NET type for each column, exports the sheet to a DataTable using full type detection (CheckMixedValueType), compares each DataColumn.DataType with the predefined types, aborts on any mismatch, and finally generates a SQL CREATE TABLE script with column‑type analysis via SqlScriptSaveOptions.
// Keywords: Aspose.Cells | C# ExportDataTable | Validate column data types | CheckMixedValueType | DataColumn DataType comparison | .NET type validation | SqlScriptSaveOptions | Generate SQL script from Excel | Excel to database import | Bulk insert type safety
// Common Searches: Aspose.Cells validate column type C# | ExportDataTable check mixed value type | compare DataColumn.DataType with expected .NET type | generate SQL CREATE TABLE from workbook Aspose.Cells | prevent type mismatch when importing Excel to SQL Server
// Developer Intent: Ensure every column exported from an Excel worksheet matches a predefined .NET type before loading the data into a database.
// Use Cases: Detect and abort on mismatched column types during bulk import | Create a reliable CREATE TABLE script that mirrors verified column types | Automate schema validation for Excel‑to‑SQL data pipelines | Log detailed type‑mismatch errors for data‑quality monitoring
// AI Prompts: Write a C# method that takes a Worksheet and a Dictionary<string, Type> of expected column types, exports the data to a DataTable with CheckMixedValueType enabled, and returns a list of columns whose DataType does not match the expected .NET type. | Generate Aspose.Cells code that exports a worksheet to a DataTable, validates each DataColumn.DataType against a predefined type map, and throws an informative exception if any column type differs. | Provide a sample that saves a workbook as a SQL script using SqlScriptSaveOptions with CheckAllDataForColumnType set to true, ensuring the generated CREATE TABLE reflects the verified column types.

using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Saving;

// This example creates a workbook, defines the expected .NET type for each column, exports the sheet to a DataTable using full type detection (CheckMixedValueType), compares each DataColumn.DataType with the predefined types, aborts on any mismatch, and finally generates a SQL CREATE TABLE script with column‑type analysis via SqlScriptSaveOptions.
class Program
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a workbook and populate it with sample data
        // -------------------------------------------------
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Header row
        ws.Cells["A1"].PutValue("ID");
        ws.Cells["B1"].PutValue("Name");
        ws.Cells["C1"].PutValue("Salary");

        // Data rows
        ws.Cells["A2"].PutValue(1);
        ws.Cells["B2"].PutValue("John");
        ws.Cells["C2"].PutValue(5000.75);

        ws.Cells["A3"].PutValue(2);
        ws.Cells["B3"].PutValue("Jane");
        ws.Cells["C3"].PutValue(6200.00);

        // -------------------------------------------------
        // 2. Define the expected .NET types for each column
        // -------------------------------------------------
        var expectedTypes = new Dictionary<string, Type>
        {
            { "ID", typeof(int) },
            { "Name", typeof(string) },
            { "Salary", typeof(double) } // Aspose.Cells maps numeric cells to Double by default
        };

        // -------------------------------------------------
        // 3. Export the worksheet to a DataTable with full type detection
        // -------------------------------------------------
        ExportTableOptions exportOptions = new ExportTableOptions
        {
            ExportColumnName = true,      // Export first row as column names
            CheckMixedValueType = true    // Examine all rows to determine column types
        };

        // Export 3 rows (header + 2 data rows) and 3 columns
        DataTable dataTable = ws.Cells.ExportDataTable(0, 0, 3, 3, exportOptions);

        // -------------------------------------------------
        // 4. Validate that each column's DataType matches the expected .NET type
        // -------------------------------------------------
        foreach (DataColumn column in dataTable.Columns)
        {
            if (expectedTypes.TryGetValue(column.ColumnName, out Type expected))
            {
                if (column.DataType != expected)
                {
                    Console.WriteLine($"Column '{column.ColumnName}' type mismatch. Expected {expected}, but got {column.DataType}.");
                    // Abort further processing because of type mismatch
                    return;
                }
            }
            else
            {
                Console.WriteLine($"No expected type defined for column '{column.ColumnName}'.");
                return;
            }
        }

        Console.WriteLine("All column types match the expected .NET types.");

        // -------------------------------------------------
        // 5. Generate a SQL script with full column‑type analysis
        // -------------------------------------------------
        SqlScriptSaveOptions sqlOptions = new SqlScriptSaveOptions
        {
            TableName = "Employees",
            CreateTable = true,
            CheckAllDataForColumnType = true // Examine all rows when determining column types
        };

        // Save the workbook as a SQL script file
        wb.Save("Employees.sql", sqlOptions);

        Console.WriteLine("SQL script generated successfully.");
    }
}
