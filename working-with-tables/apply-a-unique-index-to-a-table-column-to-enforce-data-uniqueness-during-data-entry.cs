using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsUniqueIndexDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a header row
            // Column A will be the unique identifier (primary key)
            sheet.Cells["A1"].PutValue("EmployeeID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Department");

            // Sample rows (duplicate EmployeeID will be rejected by the primary key constraint)
            sheet.Cells["A2"].PutValue(101);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue("HR");

            sheet.Cells["A3"].PutValue(102);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue("Finance");

            sheet.Cells["A4"].PutValue(101); // Duplicate ID (will violate primary key when exported)
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue("IT");

            // Configure SQL script save options
            SqlScriptSaveOptions saveOptions = new SqlScriptSaveOptions
            {
                TableName = "Employees",   // Name of the target table
                CreateTable = true,        // Generate CREATE TABLE statement
                HasHeaderRow = true,       // First row contains column names
                PrimaryKey = 0             // Set column A (index 0) as the primary key to enforce uniqueness
            };

            // Save the workbook as an SQL script file
            string outputPath = "EmployeesExport.sql";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"SQL script saved to '{outputPath}' with column 0 as the primary key.");
        }
    }
}