using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

class ApplyUniqueIndexDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header row
        sheet.Cells["A1"].PutValue("EmployeeID");
        sheet.Cells["B1"].PutValue("Name");

        // Add sample data (including a duplicate ID to illustrate uniqueness enforcement)
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");
        sheet.Cells["A4"].PutValue(1); // Duplicate ID
        sheet.Cells["B4"].PutValue("Charlie");

        // Configure SQL script save options:
        // - TableName: name of the generated table
        // - CreateTable: generate CREATE TABLE statement
        // - HasHeaderRow: first row contains column names
        // - PrimaryKey: set column index 0 (EmployeeID) as the primary key (unique index)
        SqlScriptSaveOptions saveOptions = new SqlScriptSaveOptions
        {
            TableName = "Employees",
            CreateTable = true,
            HasHeaderRow = true,
            PrimaryKey = 0
        };

        // Save the workbook as an SQL script; the script will include a PRIMARY KEY constraint on EmployeeID
        workbook.Save("Employees.sql", saveOptions);
    }
}