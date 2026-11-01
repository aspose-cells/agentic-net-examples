// Title: How to bind a List<Employee> as a data source for WorkbookDesigner smart‑marker merging in Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a List<Employee>, loads an Excel template containing ${Employees.*} markers, and uses WorkbookDesigner.SetDataSource to merge the data. | Show how to substitute the List<Employee> with a DataTable while preserving the same ${Employees.*} smart markers in the template. | Add a calculated Bonus property to the Employee class and demonstrate merging it via a new ${Employees.Bonus} smart marker.
// Common Searches: Aspose.Cells WorkbookDesigner setdata source list of objects example | C# smart markers using custom POCO collection in Excel template | how to merge employee data into Excel with Aspose.Cells smart markers | replace List<Employee> with DataTable for WorkbookDesigner in .NET | add computed field to smart marker collection in Aspose.Cells
// Tags: WorkbookDesigner SetDataSource with List<T> | Aspose.Cells smart markers custom POCO collection | Excel template merging using ${Employees.*} | C# replace List with DataTable in WorkbookDesigner | computed property merging Aspose.Cells smart markers

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCustomDataSource
{
    // Define a custom class that represents a row of data.
    // The program defines an Employee POCO, builds a List<Employee>, loads an Excel template that contains ${Employees.Name}, ${Employees.Department}, and ${Employees.Salary} smart markers, assigns the list to WorkbookDesigner via SetDataSource("Employees", employees), processes the template to merge the data, and saves the populated workbook as Result.xlsx.
    public class Employee
    {
        public string Name { get; set; }
        public string Department { get; set; }
        public double Salary { get; set; }

        public Employee(string name, string department, double salary)
        {
            Name = name;
            Department = department;
            Salary = salary;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // 1. Create a collection of custom objects.
            List<Employee> employees = new List<Employee>
            {
                new Employee("John Doe", "Finance", 75000),
                new Employee("Jane Smith", "HR", 68000),
                new Employee("Bob Johnson", "IT", 82000)
            };

            // 2. Load the Excel template that contains merge fields.
            //    The template should have placeholders like ${Employees.Name}, ${Employees.Department}, ${Employees.Salary}
            Workbook workbook = new Workbook("Template.xlsx");

            // 3. Initialize WorkbookDesigner and set the custom collection as a data source.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            // The name "Employees" will be used in the template as the collection identifier.
            designer.SetDataSource("Employees", employees);

            // 4. Process the template to merge data.
            designer.Process();

            // 5. Save the resulting workbook.
            workbook.Save("Result.xlsx");
        }
    }
}
