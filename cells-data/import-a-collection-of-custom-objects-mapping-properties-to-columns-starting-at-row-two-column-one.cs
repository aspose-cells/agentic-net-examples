// Title: Import Custom Objects into Excel Starting at A2 with Aspose.Cells (C#)
// Description: Demonstrates how to create a Workbook, define an Employee class, build a List<Employee>, and use Cells.ImportCustomObjects to map class properties to worksheet columns. The sample adds a header row, starts the import at cell A2, formats dates as yyyy‑MM‑dd, converts numeric strings, and saves the file as EmployeesImport.xlsx.
// Keywords: Aspose.Cells ImportCustomObjects C# | map object properties to Excel columns | import list of objects starting at specific row | export C# collection to Excel | date formatting Aspose.Cells | convert string to number Aspose.Cells
// Common Searches: Aspose.Cells import list of objects starting at A2 | C# ImportCustomObjects header row example | how to map class properties to Excel columns Aspose.Cells | set date format when importing objects with Aspose.Cells | convert numeric strings during Excel export C#
// Developer Intent: Load a collection of Employee objects into a worksheet, map each property to a column, include headers, and begin writing at cell A2.
// Use Cases: Generate payroll spreadsheets with employee IDs, names, salaries, and hire dates formatted uniformly. | Populate a pre‑designed report template by inserting rows below existing headers and mapping data automatically. | Append new records to an existing Excel file while preserving column order and applying consistent date formatting.
// AI Prompts: Show how to import the same collection starting at cell B3 without adding a header row. | Provide code that changes the property order and uses a custom date format like "dd MMM yyyy". | Explain performance tips for importing large collections, such as disabling row insertion and handling formatting manually.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportDemo
{
    // Define a custom data class whose properties will be mapped to worksheet columns
    // Demonstrates how to create a Workbook, define an Employee class, build a List<Employee>, and use Cells.ImportCustomObjects to map class properties to worksheet columns. The sample adds a header row, starts the import at cell A2, formats dates as yyyy‑MM‑dd, converts numeric strings, and saves the file as EmployeesImport.xlsx.
    public class Employee
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public decimal Salary { get; set; }
        public DateTime HireDate { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Prepare a collection of custom objects
            List<Employee> employees = new List<Employee>
            {
                new Employee { Id = 1, Name = "Alice", Salary = 75000m, HireDate = new DateTime(2020, 5, 10) },
                new Employee { Id = 2, Name = "Bob", Salary = 62000m, HireDate = new DateTime(2019, 3, 22) },
                new Employee { Id = 3, Name = "Charlie", Salary = 88000m, HireDate = new DateTime(2021, 11, 1) }
            };

            // Specify the property names to map to columns (order matters)
            string[] propertyNames = { "Id", "Name", "Salary", "HireDate" };

            // Import the collection starting at row 2 (index 1) and column 1 (index 0)
            // Parameters:
            //   list                : employees collection
            //   propertyNames       : columns mapping
            //   isPropertyNameShown : include header row
            //   firstRow            : 1 (second row)
            //   firstColumn         : 0 (first column)
            //   rowNumber           : number of rows to import (employees.Count)
            //   insertRows          : true (add rows if needed)
            //   dateFormatString    : desired date format
            //   convertStringToNumber: true (convert numeric strings)
            int importedRows = cells.ImportCustomObjects(
                employees,
                propertyNames,
                true,
                1,
                0,
                employees.Count,
                true,
                "yyyy-MM-dd",
                true
            );

            Console.WriteLine($"Imported {importedRows} rows starting at cell A2.");

            // Save the workbook to a file (lifecycle save)
            workbook.Save("EmployeesImport.xlsx");
        }
    }
}
