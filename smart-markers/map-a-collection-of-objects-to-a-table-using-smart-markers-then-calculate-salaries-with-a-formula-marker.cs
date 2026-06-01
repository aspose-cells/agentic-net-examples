using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerSalaryDemo
{
    // Sample data class
    public class Employee
    {
        public string Name { get; set; }
        public double Hours { get; set; }
        public double Rate { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Set up the template with smart markers
            // Header row
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Hours");
            cells["C1"].PutValue("Rate");
            cells["D1"].PutValue("Salary");

            // Row that will be repeated for each Employee object
            // Smart markers reference the data source name "Employees"
            cells["A2"].PutValue("&=$Employees.Name");
            cells["B2"].PutValue("&=$Employees.Hours");
            cells["C2"].PutValue("&=$Employees.Rate");
            // Formula marker: Salary = Hours * Rate
            cells["D2"].Formula = "=B2*C2";

            // 3. Prepare a collection of Employee objects
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "Alice", Hours = 40, Rate = 25 },
                new Employee { Name = "Bob",   Hours = 35, Rate = 30 },
                new Employee { Name = "Carol", Hours = 45, Rate = 22 }
            };

            // 4. Create a WorkbookDesigner, assign the workbook, and set the data source
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            // Bind the collection to the name used in smart markers
            designer.SetDataSource("Employees", employees);
            // Ensure that formulas (Salary) are repeated for each generated row
            designer.RepeatFormulasWithSubtotal = true;

            // 5. Process the smart markers to populate the table
            designer.Process();

            // 6. Calculate all formulas so that Salary values are evaluated
            workbook.CalculateFormula();

            // 7. Save the resulting workbook
            workbook.Save("EmployeesSalary.xlsx");
        }
    }
}