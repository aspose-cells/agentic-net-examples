// Title: Use WorkbookDesigner.LineByLine to merge a simple list with smart markers in Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# console application that creates a Workbook, places smart markers for employee Name, Age, and Department, binds a List<Employee> to the marker "Employees", enables line‑by‑line mode, and saves the file as an .xlsx document. | Show how setting WorkbookDesigner.LineByLine = true causes each item in a collection to be written to a separate row when processing smart markers with Aspose.Cells. | Provide code that binds a custom object collection to a smart marker and generates an Excel sheet using line‑by‑line smart marker expansion via WorkbookDesigner.
// Common Searches: asp.net aspose.cells workbookdesigner linebyline example c# | how to export a List<Employee> to Excel using smart markers line by line | set linebyline true for master smart markers in Aspose.Cells | merge simple list into Excel template with line‑by‑line smart markers c#
// Tags: WorkbookDesigner line-by-line mode | Aspose.Cells smart markers list binding | C# export object collection to Excel | Excel generation with smart markers | line-by-line smart marker expansion

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsLineByLineDemo
{
    // Simple data class representing an employee
    // The sample creates a new workbook, inserts smart markers for employee Name, Age, and Department, binds a List<Employee> to the "Employees" marker, enables WorkbookDesigner.LineByLine for line‑by‑line processing, processes the markers, and saves the populated worksheet as LineByLineOutput.xlsx.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (template) and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert smart markers for a simple list (master markers)
            // These markers will be processed line by line
            sheet.Cells["A1"].PutValue("&Employees.Name");
            sheet.Cells["B1"].PutValue("&Employees.Age");
            sheet.Cells["C1"].PutValue("&Employees.Department");

            // Prepare sample data source (a list of employees)
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Age = 35, Department = "Sales" },
                new Employee { Name = "Jane Smith", Age = 28, Department = "Marketing" },
                new Employee { Name = "Bob Johnson", Age = 42, Department = "HR" }
            };

            // Initialize the WorkbookDesigner with the workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set LineByLine to true to ensure processing line by line
            designer.LineByLine = true;

            // Bind the data source to the smart marker name "Employees"
            designer.SetDataSource("Employees", employees);

            // Process the smart markers and populate the worksheet
            designer.Process();

            // Save the resulting workbook
            workbook.Save("LineByLineOutput.xlsx", SaveFormat.Xlsx);
        }
    }
}
