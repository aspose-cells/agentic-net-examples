// Title: Generate an Excel employee salary report from a C# List using Aspose.Cells smart markers and a formula marker
// Common Searches: C# Aspose.Cells populate Excel rows from List<Employee> using smart markers | How to calculate total salary column with a formula marker in Aspose.Cells WorkbookDesigner | Aspose.Cells example for generating employee salary table with smart markers and formulas | Binding a collection to an Excel worksheet and evaluating formulas with WorkbookDesigner in C# | Create an XLSX employee report with base salary, bonus, and total using smart markers
// Tags: smart markers collection binding Aspose.Cells | formula marker total salary calculation Aspose.Cells | WorkbookDesigner data source binding C# | export employee report to XLSX Aspose.Cells | dynamic row generation with smart markers

// Import Aspose.Cells namespace
using Aspose.Cells;
using System;
using System.Collections.Generic;

// Define a simple Employee class
// The sample creates a workbook, adds header cells, inserts smart markers for Name, BaseSalary, and Bonus, places a formula marker to compute Total Salary, binds a List<Employee> as the "Employees" data source, processes the markers with WorkbookDesigner to generate rows and evaluate the formula, and saves the result as EmployeesReport.xlsx.
public class Employee
{
    public string Name { get; set; }
    public double BaseSalary { get; set; }
    public double Bonus { get; set; }
}

// Main execution
public class Program
{
    public static void Main()
    {
        // 1. Create a new workbook
        Workbook workbook = new Workbook();

        // 2. Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // 3. Set header titles
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Base Salary");
        sheet.Cells["C1"].PutValue("Bonus");
        sheet.Cells["D1"].PutValue("Total Salary");

        // 4. Insert smart markers for the data rows (starting at row 2)
        sheet.Cells["A2"].PutValue("&=Employees.Name");
        sheet.Cells["B2"].PutValue("&=Employees.BaseSalary");
        sheet.Cells["C2"].PutValue("&=Employees.Bonus");
        // 5. Insert a formula marker that will calculate total salary for each row
        //    The formula uses relative references; Aspose.Cells will adjust it for each populated row.
        sheet.Cells["D2"].Formula = "=B2+C2";

        // 6. Prepare a collection of Employee objects
        List<Employee> employeeList = new List<Employee>()
        {
            new Employee() { Name = "Alice", BaseSalary = 5000, Bonus = 500 },
            new Employee() { Name = "Bob",   BaseSalary = 4500, Bonus = 300 },
            new Employee() { Name = "Carol", BaseSalary = 6000, Bonus = 800 }
        };

        // 7. Use WorkbookDesigner to process smart markers
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource("Employees", employeeList);
        designer.Process(); // Populates the table and evaluates the formula marker

        // 8. Save the resulting workbook to a file
        workbook.Save("EmployeesReport.xlsx", SaveFormat.Xlsx);
    }
}
