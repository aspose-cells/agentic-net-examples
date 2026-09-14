// Title: Import a List<Employee> into an Excel worksheet starting at cell A2 using Aspose.Cells for .NET
// AI Prompts: Write C# code that calls Worksheet.Cells.ImportCustomObjects to load a List<Employee> into the first worksheet beginning at row 2, column 1, mapping Id, Name, Salary, and HireDate to columns A‑D. | Show how to apply a custom date format (yyyy‑MM‑dd) and enable string‑to‑number conversion while importing objects with Aspose.Cells. | Demonstrate adding a header row that displays the property names before the data rows when importing custom objects.
// Common Searches: Aspose.Cells C# import custom object list starting at A2 without header row | Worksheet.Cells.ImportCustomObjects example with row and column offset | How to map Employee class properties to Excel columns using Aspose.Cells | Set date format for DateTime values during ImportCustomObjects in Aspose.Cells | Enable automatic numeric conversion when importing string values with Aspose.Cells .NET
// Tags: import custom objects Aspose.Cells C# | ImportCustomObjects start row column offset | map Employee properties to Excel columns | custom date format Aspose.Cells import | auto insert rows Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportCustomObjectsDemo
{
    // Define a custom data class whose properties will be mapped to worksheet columns
    // C# example that creates a Workbook, prepares a List<Employee>, and uses Worksheet.Cells.ImportCustomObjects to write the Id, Name, Salary, and HireDate fields into columns A‑D starting at cell A2 (row 2, column 1) without a header row. The import applies the "yyyy-MM-dd" date format, converts strings to numbers, inserts rows as needed, and saves the result as EmployeesImport.xlsx.
    public class Employee
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public decimal Salary { get; set; }
        public DateTime HireDate { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare a collection of custom objects to import
            List<Employee> employees = new List<Employee>
            {
                new Employee { Id = 1, Name = "Alice", Salary = 75000m, HireDate = new DateTime(2020, 5, 10) },
                new Employee { Id = 2, Name = "Bob",   Salary = 62000m, HireDate = new DateTime(2019, 3, 22) },
                new Employee { Id = 3, Name = "Carol", Salary = 88000m, HireDate = new DateTime(2021, 11, 1) }
            };

            // Specify the property names to map to columns (order matters)
            string[] propertyNames = { "Id", "Name", "Salary", "HireDate" };

            // Import the collection starting at row 2 (index 1) and column 1 (index 0)
            // Parameters:
            //   list                : employees collection
            //   propertyNames       : columns mapping
            //   isPropertyNameShown : false (no header row)
            //   firstRow            : 1 (second row)
            //   firstColumn         : 0 (first column)
            //   rowNumber           : employees.Count
            //   insertRows          : true (add rows if needed)
            //   dateFormatString    : "yyyy-MM-dd"
            //   convertStringToNumber: true
            int importedRows = worksheet.Cells.ImportCustomObjects(
                employees,
                propertyNames,
                false,
                1,
                0,
                employees.Count,
                true,
                "yyyy-MM-dd",
                true
            );

            Console.WriteLine($"Imported {importedRows} rows starting at cell A2.");

            // Save the workbook to a file
            workbook.Save("EmployeesImport.xlsx");
        }
    }
}
