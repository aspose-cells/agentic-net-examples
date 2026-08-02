using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerSalaryDemo
{
    // Sample data class
    public class Employee
    {
        public string Name { get; set; }
        public double Salary { get; set; }
        public double Bonus { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Build a simple template with smart markers
            // Header row
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Salary");
            cells["C1"].PutValue("Bonus");
            cells["D1"].PutValue("Total");

            // Data row with smart markers
            // &=$Employees.Name will be replaced by each employee's Name
            // &=$Employees.Salary will be replaced by each employee's Salary
            // &=$Employees.Bonus will be replaced by each employee's Bonus
            cells["A2"].PutValue("&=$Employees.Name");
            cells["B2"].PutValue("&=$Employees.Salary");
            cells["C2"].PutValue("&=$Employees.Bonus");
            // Formula marker to calculate total compensation
            cells["D2"].Formula = "=B2+C2";

            // 3. Prepare a collection of objects
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Salary = 50000, Bonus = 5000 },
                new Employee { Name = "Jane Smith", Salary = 62000, Bonus = 6200 },
                new Employee { Name = "Bob Johnson", Salary = 47000, Bonus = 4700 }
            };

            // 4. Set the collection as a data source for the smart marker named "Employees"
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Employees", employees);

            // 5. Process the smart markers (lifecycle rule: process)
            designer.Process();

            // 6. Calculate all formulas, including the Total column (lifecycle rule: calculate)
            workbook.CalculateFormula();

            // 7. Save the result (lifecycle rule: save)
            workbook.Save("EmployeesSalaryReport.xlsx");
        }
    }
}