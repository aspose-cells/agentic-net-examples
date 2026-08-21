// Title: C# – Populate Excel with Smart Markers from a List and Compute Total Salary via Formula Marker
// Description: Demonstrates how to bind a List<Employee> to an Excel worksheet using Aspose.Cells smart markers, insert a formula marker to calculate Total Salary (BaseSalary + BaseSalary*BonusRate), evaluate all formulas, and save the workbook as EmployeesSalaryReport.xlsx.
// Keywords: Aspose.Cells smart markers C# | WorkbookDesigner SetDataSource example | populate Excel table from List | formula marker calculate salary | Excel automation payroll report | C# calculate Excel formulas Aspose | dynamic salary sheet Aspose.Cells | Excel export employee data .NET
// Common Searches: Aspose.Cells smart markers bind collection C# | how to use formula marker in Aspose.Cells | calculate total salary with smart markers | process smart markers then evaluate formulas | C# export payroll to Excel using Aspose
// Developer Intent: Generate an Excel payroll sheet by mapping a collection of Employee objects with smart markers and automatically computing each employee's total compensation using a formula marker.
// Use Cases: Create a payroll report that expands rows for each employee and calculates total compensation on the fly. | Build a dynamic salary worksheet where bonus percentages are applied via a single formula marker. | Export employee records to Excel with automatic currency formatting after formula evaluation.
// AI Prompts: Show how to modify the formula marker to subtract tax deductions from the total salary. | Provide an example of using smart markers with nested collections to produce department‑wise salary summaries. | Explain how to apply currency formatting to the Total Salary column after calling CalculateFormula.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerExample
{
    // Sample data class
    // Demonstrates how to bind a List<Employee> to an Excel worksheet using Aspose.Cells smart markers, insert a formula marker to calculate Total Salary (BaseSalary + BaseSalary*BonusRate), evaluate all formulas, and save the workbook as EmployeesSalaryReport.xlsx.
    public class Employee
    {
        public string Name { get; set; }
        public double BaseSalary { get; set; }
        public double BonusRate { get; set; }   // e.g., 0.10 for 10%
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare sample data
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", BaseSalary = 50000, BonusRate = 0.10 },
                new Employee { Name = "Jane Smith", BaseSalary = 62000, BonusRate = 0.15 },
                new Employee { Name = "Bob Johnson", BaseSalary = 48000, BonusRate = 0.08 }
            };

            // 2. Create a new workbook (template) and define smart markers
            Workbook workbook = new Workbook();                     // create
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Base Salary");
            cells["C1"].PutValue("Bonus Rate");
            cells["D1"].PutValue("Total Salary");

            // Smart markers start at row 2
            cells["A2"].PutValue("&=$Employees.Name");
            cells["B2"].PutValue("&=$Employees.BaseSalary");
            cells["C2"].PutValue("&=$Employees.BonusRate");
            // Formula marker: Total Salary = BaseSalary + (BaseSalary * BonusRate)
            cells["D2"].PutValue("=B2 + (B2 * C2)");               // formula marker

            // 3. Set up WorkbookDesigner and bind the data source
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Employees", employees);        // bind collection

            // 4. Process smart markers (populate the table)
            designer.Process();                                    // process all sheets

            // 5. Calculate formulas so that Total Salary column is evaluated
            workbook.CalculateFormula();                           // calculate all formulas

            // 6. Save the result
            workbook.Save("EmployeesSalaryReport.xlsx");           // save
        }
    }
}
