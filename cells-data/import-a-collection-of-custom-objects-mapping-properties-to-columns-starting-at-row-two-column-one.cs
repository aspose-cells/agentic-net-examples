using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportDemo
{
    // Sample custom object
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Prepare a collection of custom objects
            List<Employee> employees = new List<Employee>
            {
                new Employee { Id = 1, Name = "Alice", Salary = 75000m, HireDate = new DateTime(2020, 5, 10) },
                new Employee { Id = 2, Name = "Bob",   Salary = 62000m, HireDate = new DateTime(2019, 3, 22) },
                new Employee { Id = 3, Name = "Carol", Salary = 88000m, HireDate = new DateTime(2021, 11, 1) }
            };

            // Define the property names to map to columns (order matters)
            string[] propertyNames = { "Id", "Name", "Salary", "HireDate" };

            // Import the collection starting at row 2 (index 1) and column 1 (index 0)
            // Parameters:
            //   list                : employees
            //   propertyNames       : propertyNames
            //   isPropertyNameShown : true (adds header row)
            //   firstRow            : 1  (second row in the sheet)
            //   firstColumn         : 0  (first column)
            //   rowNumber           : employees.Count
            //   insertRows          : true (adds rows if needed)
            //   dateFormatString    : "yyyy-MM-dd"
            //   convertStringToNumber: true
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

            Console.WriteLine($"Imported {importedRows} rows starting at B2.");

            // Save the workbook
            workbook.Save("EmployeesImport.xlsx");
        }
    }
}