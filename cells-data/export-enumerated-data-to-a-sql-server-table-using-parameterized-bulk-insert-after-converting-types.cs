// Title: Export Enum Data to SQL Server with Aspose.Cells Bulk INSERT and Custom Column Types
// Description: Creates a workbook, writes enum status codes to two columns, defines a CustomSqlColumnTypeMap that maps integers to SqlInteger and strings to SqlVarChar, configures SqlScriptSaveOptions to generate a CREATE TABLE statement and bulk INSERT script for dbo.StatusLookup, and saves the result as a .sql file ready for execution on SQL Server.
// Keywords: Aspose.Cells | C# | SqlScriptSaveOptions | custom SqlScriptColumnTypeMap | bulk INSERT | SQL Server | enum export | generate CREATE TABLE script | Excel to SQL script | parameterized bulk insert
// Common Searches: Aspose.Cells export enum to SQL Server | How to map Excel columns to SQL data types with Aspose.Cells | Generate bulk INSERT script from workbook C# | Custom column type mapping Aspose.Cells SqlScriptSaveOptions | Create lookup table from enum using Aspose.Cells
// Developer Intent: Produce a SQL script that creates a table and inserts enum values from an Aspose.Cells workbook, using a custom type map for accurate SQL data types.
// Use Cases: Automatically generate a status‑code lookup table in SQL Server from a C# enum. | Create reusable scripts that convert any enumerated dataset into CREATE TABLE and INSERT statements. | Control the SQL data types of exported columns (e.g., INTEGER, VARCHAR) by extending SqlScriptColumnTypeMap.
// AI Prompts: Write C# code that uses Aspose.Cells to export a DataTable to a SQL script with a custom SqlScriptColumnTypeMap that maps DateTime columns to SqlDateTime. | Show how to modify CustomSqlColumnTypeMap to return SqlDecimal for decimal fields and SqlVarChar(100) for string fields. | Explain how to run the generated StatusLookup.sql file with sqlcmd or PowerShell to perform a bulk insert into SQL Server.

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Saving;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsSqlExport
{
    // Custom column type map that uses SqlDataType enum to define SQL types
    // Creates a workbook, writes enum status codes to two columns, defines a CustomSqlColumnTypeMap that maps integers to SqlInteger and strings to SqlVarChar, configures SqlScriptSaveOptions to generate a CREATE TABLE statement and bulk INSERT script for dbo.StatusLookup, and saves the result as a .sql file ready for execution on SQL Server.
    public class CustomSqlColumnTypeMap : SqlScriptColumnTypeMap
    {
        // Return the SQL type for numeric columns (e.g., INT)
        public override string GetNumbericType()
        {
            // Use SqlDataType enum to illustrate the intended type
            // Here we map to SQL INTEGER
            return SqlDataType.SqlInteger.ToString();
        }

        // Return the SQL type for string columns (e.g., VARCHAR(255))
        public override string GetStringType()
        {
            // Map to SQL VARCHAR
            return SqlDataType.SqlVarChar.ToString();
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Prepare enumerated data (example enum)
            // Assume we have an enum representing status codes
            // We'll write the enum names and their integer values into the sheet
            sheet.Cells["A1"].PutValue("StatusCode");   // Header
            sheet.Cells["B1"].PutValue("Description"); // Header

            // Sample enum values
            var statuses = new (int Code, string Description)[]
            {
                (0, "Success"),
                (1, "Warning"),
                (2, "Error"),
                (3, "Critical")
            };

            // Populate the worksheet starting from row 2 (index 1)
            for (int i = 0; i < statuses.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(statuses[i].Code);        // Column A
                sheet.Cells[i + 1, 1].PutValue(statuses[i].Description); // Column B
            }

            // 3. Configure SQL script save options for bulk INSERT
            SqlScriptSaveOptions sqlOptions = new SqlScriptSaveOptions
            {
                TableName = "dbo.StatusLookup",          // Target SQL Server table
                CreateTable = true,                      // Generate CREATE TABLE statement
                OperatorType = SqlScriptOperatorType.Insert, // Use INSERT statements
                HasHeaderRow = true,                     // First row contains column names
                ExportAsString = false,                  // Export values with proper types
                CheckAllDataForColumnType = true,        // Examine all rows to infer column types
                ColumnTypeMap = new CustomSqlColumnTypeMap() // Use custom type mapping
            };

            // 4. Save the workbook as a SQL script file
            // The generated script can be executed with SQL Server's bulk insert utilities
            workbook.Save("StatusLookup.sql", sqlOptions);

            // Optional: display the generated script content
            Console.WriteLine("SQL script generated:");
            Console.WriteLine(System.IO.File.ReadAllText("StatusLookup.sql"));
        }
    }
}
