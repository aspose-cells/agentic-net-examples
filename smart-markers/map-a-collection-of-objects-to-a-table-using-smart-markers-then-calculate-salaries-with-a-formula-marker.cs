// Title: Create an Excel employee report with Aspose.Cells smart markers and a formula marker to compute total compensation in C#
// AI Prompts: Write C# code that binds a List<Employee> to WorkbookDesigner, inserts smart markers for Name, Position, Salary, and Bonus, adds a formula smart marker to calculate Salary+Bonus, processes the markers, and saves the workbook as an .xlsx file. | Show how to combine Aspose.Cells smart markers and a formula marker to generate a table and a computed Total Compensation column from a collection of objects.
// Common Searches: how to use Aspose.Cells smart markers with a List<T> in C# | Aspose.Cells formula smart marker calculate column value | generate Excel report from collection using WorkbookDesigner C# | smart markers total compensation column Aspose.Cells example | binding employee list to Excel template Aspose.Cells
// Tags: Aspose.Cells WorkbookDesigner smart markers C# | smart marker formula marker total compensation | populate Excel table from List<Employee> Aspose | calculate column value with formula marker Aspose.Cells | export employee data to XLSX using smart markers

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerExample
{
    // Simple POCO representing an employee
    // // Generates EmployeesReport.xlsx containing a table of employee Name, Position, Salary, Bonus, and a computed Total Compensation column using a formula smart marker, by binding a List<Employee> to WorkbookDesigner and processing the smart markers.
    public class Employee
    {
        public string Name { get; set; }
        public string Position { get; set; }
        public double Salary { get; set; }
        public double Bonus { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Prepare sample data
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Position = "Developer", Salary = 60000, Bonus = 5000 },
                new Employee { Name = "Jane Smith", Position = "Designer", Salary = 55000, Bonus = 4000 },
                new Employee { Name = "Bob Johnson", Position = "Manager", Salary = 80000, Bonus = 10000 }
            };

            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Employees";

            // Write table headers
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Position");
            sheet.Cells["C1"].PutValue("Salary");
            sheet.Cells["D1"].PutValue("Bonus");
            sheet.Cells["E1"].PutValue("Total Compensation");

            // Insert smart markers for the data rows (starting at row 2)
            // Smart markers are enclosed in &= & and reference the collection name "Employees"
            sheet.Cells["A2"].PutValue("&=Employees.Name&");
            sheet.Cells["B2"].PutValue("&=Employees.Position&");
            sheet.Cells["C2"].PutValue("&=Employees.Salary&");
            sheet.Cells["D2"].PutValue("&=Employees.Bonus&");
            // Formula marker to calculate Salary + Bonus for each employee
            sheet.Cells["E2"].PutValue("&=Employees.Salary+Employees.Bonus&");

            // Set the smart marker range (optional, Aspose.Cells can infer it)
            // Here we let the designer process the whole sheet.

            // Bind the data source to the workbook designer
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Employees", employees);
            designer.Process(); // This replaces smart markers with actual data and evaluates formulas

            // Save the result
            workbook.Save("EmployeesReport.xlsx", SaveFormat.Xlsx);
        }
    }
}
