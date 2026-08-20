// Title: C# – Export Excel column as Primary Key with Aspose.Cells SqlScriptSaveOptions
// Description: Demonstrates how to create a workbook, populate it with sample employee data, and use Aspose.Cells SqlScriptSaveOptions to generate an SQL script that defines a table named **Employees** with the first column (EmployeeID) set as a PRIMARY KEY, ensuring uniqueness during data entry.
// Keywords: Aspose.Cells | C# | SqlScriptSaveOptions | primary key | unique index | export Excel to SQL | CREATE TABLE script | data uniqueness | Excel to database migration
// Common Searches: Aspose.Cells set primary key when saving to SQL script | C# export Excel column as unique index | SqlScriptSaveOptions PrimaryKey example | generate CREATE TABLE from Excel with Aspose.Cells | prevent duplicate IDs in exported SQL script
// Developer Intent: Produce an SQL script from an Excel worksheet where a chosen column is declared as the primary key to enforce uniqueness.
// Use Cases: Migrating an employee roster from Excel to a relational database while guaranteeing that EmployeeID values are unique. | Automating schema generation for data‑import pipelines, creating tables with built‑in primary‑key constraints directly from spreadsheets. | Building data‑validation tools that export Excel data to SQL scripts and automatically reject duplicate identifier rows.
// AI Prompts: Generate C# code that uses Aspose.Cells to export a worksheet to an SQL script with a composite primary key on columns A and C. | Explain how to validate and remove duplicate rows in a worksheet before applying SqlScriptSaveOptions. | Show how to configure SqlScriptSaveOptions to add a UNIQUE constraint instead of a primary key for a non‑identifier column.

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

// Demonstrates how to create a workbook, populate it with sample employee data, and use Aspose.Cells SqlScriptSaveOptions to generate an SQL script that defines a table named **Employees** with the first column (EmployeeID) set as a PRIMARY KEY, ensuring uniqueness during data entry.
class UniqueIndexDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the worksheet with a header row and some sample data
        worksheet.Cells["A1"].PutValue("EmployeeID"); // Column to be made unique
        worksheet.Cells["B1"].PutValue("Name");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("Alice");
        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue("Bob");
        worksheet.Cells["A4"].PutValue(1); // Duplicate ID to illustrate uniqueness enforcement

        // Configure SQL script save options:
        // - TableName: name of the generated table
        // - PrimaryKey: index of the column that will become the primary key (unique index)
        // - CreateTable: generate CREATE TABLE statement
        // - HasHeaderRow: treat first row as column headers
        // - CheckIfTableExists: optional, set to false for this demo
        SqlScriptSaveOptions saveOptions = new SqlScriptSaveOptions
        {
            TableName = "Employees",
            PrimaryKey = 0, // Column A (EmployeeID) will be the unique index
            CreateTable = true,
            HasHeaderRow = true,
            CheckIfTableExists = false
        };

        // Save the workbook as an SQL script; the script will contain a PRIMARY KEY constraint
        workbook.Save("Employees.sql", saveOptions);
    }
}
