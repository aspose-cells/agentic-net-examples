// Title: Use Aspose.Cells smart markers in C# to populate an Excel table from a List<Employee> and calculate Salary with a formula marker
// AI Prompts: Generate C# code that creates a workbook, defines column headers, and uses Aspose.Cells smart markers to insert rows from a List<Employee>. | Write C# that binds a List<Employee> to WorkbookDesigner, processes the smart markers, and evaluates a formula marker to compute the Salary column for each row. | Show how to append a total row that sums the Salary column after smart marker processing using Aspose.Cells.
// Common Searches: aspocells c# smart markers map list of objects to excel rows | how to calculate column values with formula marker in Aspose.Cells | WorkbookDesigner set data source and evaluate formulas example | c# generate employee salary report using Aspose.Cells smart markers | aspocells formula marker BaseSalary+Bonus per row
// Tags: Aspose.Cells smart markers collection mapping | C# WorkbookDesigner bind list data source | Aspose.Cells formula marker calculate column | Excel salary report generation with Aspose.Cells | evaluate formulas after smart marker processing

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates creating a workbook, adding headers, using smart markers to map a List<Employee> to rows, applying a formula marker to compute Salary, processing with WorkbookDesigner, calculating formulas, and saving the Excel file.
class Program
{
    static void Main()
    {
        // 1. Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // 2. Define header row
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("BaseSalary");
        cells["C1"].PutValue("Bonus");
        cells["D1"].PutValue("Salary");

        // 3. Template row with smart markers
        //    &=$Employees.Property maps the property value for each object in the collection
        cells["A2"].PutValue("&=$Employees.Name");
        cells["B2"].PutValue("&=$Employees.BaseSalary");
        cells["C2"].PutValue("&=$Employees.Bonus");
        //    Formula marker: &=[ColumnNames] will be converted to a proper Excel formula per row
        cells["D2"].PutValue("&=BaseSalary+Bonus");

        // 4. Prepare a collection of objects to be mapped
        List<Employee> employees = new List<Employee>
        {
            new Employee { Name = "John", BaseSalary = 5000, Bonus = 500 },
            new Employee { Name = "Jane", BaseSalary = 6000, Bonus = 600 },
            new Employee { Name = "Bob",  BaseSalary = 5500, Bonus = 550 }
        };

        // 5. Set up WorkbookDesigner, bind the data source and process smart markers
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource("Employees", employees);
        designer.Process(); // processes all smart markers in the workbook

        // 6. Calculate formulas so that the Salary column gets evaluated
        workbook.CalculateFormula();

        // 7. Save the resulting workbook (lifecycle: save)
        workbook.Save("EmployeesSalary.xlsx");
    }

    // Simple POCO representing an employee
    public class Employee
    {
        public string Name { get; set; }
        public double BaseSalary { get; set; }
        public double Bonus { get; set; }
    }
}
