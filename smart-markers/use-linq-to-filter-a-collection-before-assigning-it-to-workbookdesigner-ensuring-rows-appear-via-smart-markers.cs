using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerLinqDemo
{
    // Simple data class representing an employee
    public class Employee
    {
        public string Name { get; set; }
        public string Department { get; set; }
        public decimal Salary { get; set; }

        public Employee(string name, string department, decimal salary)
        {
            Name = name;
            Department = department;
            Salary = salary;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare a collection of employees
            List<Employee> allEmployees = new List<Employee>
            {
                new Employee("John Doe", "Sales", 55000m),
                new Employee("Jane Smith", "HR", 48000m),
                new Employee("Bob Johnson", "Sales", 62000m),
                new Employee("Alice Brown", "IT", 72000m),
                new Employee("Tom Clark", "Sales", 50000m)
            };

            // 2. Use LINQ to filter only Sales department employees with Salary > 50000
            var filteredEmployees = allEmployees
                .Where(e => e.Department == "Sales" && e.Salary > 50000m)
                .ToList();

            // 3. Create a workbook and place smart markers for the data source
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Header row
            ws.Cells["A1"].PutValue("Name");
            ws.Cells["B1"].PutValue("Department");
            ws.Cells["C1"].PutValue("Salary");

            // Smart marker row – the marker name "Employees" will be used as the data source name
            ws.Cells["A2"].PutValue("&=Employees.Name");
            ws.Cells["B2"].PutValue("&=Employees.Department");
            ws.Cells["C2"].PutValue("&=Employees.Salary");

            // 4. Initialize WorkbookDesigner, assign the filtered collection as data source, and process
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = wb
            };

            // Bind the filtered list to the smart marker name "Employees"
            designer.SetDataSource("Employees", filteredEmployees);

            // Process the smart markers – rows will be generated for each item in filteredEmployees
            designer.Process();

            // 5. Save the resulting workbook
            wb.Save("FilteredEmployees.xlsx");
        }
    }
}