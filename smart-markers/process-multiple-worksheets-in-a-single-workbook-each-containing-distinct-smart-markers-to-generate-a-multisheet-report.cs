// Title: Generate a multi‑sheet Excel report with separate smart markers for Employees and Departments using Aspose.Cells in C#
// AI Prompts: Create a C# console program that builds an Aspose.Cells Workbook, adds worksheets named "Employees" and "Departments", places smart markers such as "&=Employees.Name" and "&=Departments.DeptName", binds List<Employee> and List<Department> objects with WorkbookDesigner, processes all markers, and saves the result as MultiSheetReport.xlsx. | Extend the solution by adding a third worksheet called "Projects", define smart markers for ProjectName and Budget, bind a List<Project> data source, and regenerate the multi‑sheet report.
// Common Searches: asp.net how to apply Aspose.Cells smart markers on multiple worksheets in one workbook | c# generate Excel file with employee and department tables using smart markers Aspose.Cells | bind separate list objects to smart markers in different sheets with WorkbookDesigner | process all smart markers in a workbook using Aspose.Cells C# example
// Tags: Aspose.Cells WorkbookDesigner process smart markers | C# generate multi‑sheet Excel with smart markers | bind List data source to smart markers Aspose.Cells | create Excel report with employee and department worksheets | smart markers across multiple worksheets Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Demonstrates building a workbook with two worksheets, inserting smart markers for Employees and Departments, binding List<Employee> and List<Department> data sources via WorkbookDesigner, processing all markers across sheets, and saving the output as MultiSheetReport.xlsx.
public class MultiSheetSmartMarkerReport
{
    public static void Main()
    {
        try
        {
            Run();
            Console.WriteLine("Report generated successfully.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error generating report: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook that will serve as the template
        Workbook workbook = new Workbook();

        // ---------- Worksheet 1 : Employees ----------
        Worksheet wsEmployees = workbook.Worksheets[0];
        wsEmployees.Name = "Employees";

        // Header row
        wsEmployees.Cells["A1"].PutValue("Name");
        wsEmployees.Cells["B1"].PutValue("Age");

        // Smart markers – they will be replaced by the data source values
        wsEmployees.Cells["A2"].PutValue("&=Employees.Name");
        wsEmployees.Cells["B2"].PutValue("&=Employees.Age");

        // ---------- Worksheet 2 : Departments ----------
        Worksheet wsDepartments = workbook.Worksheets.Add("Departments");

        // Header row
        wsDepartments.Cells["A1"].PutValue("Dept");
        wsDepartments.Cells["B1"].PutValue("Location");

        // Smart markers for the second sheet
        wsDepartments.Cells["A2"].PutValue("&=Departments.DeptName");
        wsDepartments.Cells["B2"].PutValue("&=Departments.Location");

        // ---------- Prepare data sources ----------
        List<Employee> employees = new List<Employee>
        {
            new Employee { Name = "John Doe", Age = 30 },
            new Employee { Name = "Jane Smith", Age = 28 }
        };

        List<Department> departments = new List<Department>
        {
            new Department { DeptName = "HR", Location = "New York" },
            new Department { DeptName = "IT", Location = "San Francisco" }
        };

        // ---------- Initialize WorkbookDesigner ----------
        // Use the constructor that accepts a Workbook instance
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Bind the data sources to the names referenced in the smart markers
        designer.SetDataSource("Employees", employees);
        designer.SetDataSource("Departments", departments);

        // Process all smart markers across all worksheets
        designer.Process();

        // ---------- Save the generated multi‑sheet report ----------
        string outputPath = "MultiSheetReport.xlsx";

        // Ensure the directory exists before saving
        string? directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        workbook.Save(outputPath);
    }

    // Simple POCO classes representing the data structures
    public class Employee
    {
        public string? Name { get; set; }
        public int Age { get; set; }
    }

    public class Department
    {
        public string? DeptName { get; set; }
        public string? Location { get; set; }
    }
}
