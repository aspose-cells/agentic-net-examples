// Title: Aspose.Cells for .NET: Use WorkbookDesigner.LineByLine to merge a list with smart markers (C#)
// Description: Demonstrates enabling WorkbookDesigner.LineByLine in C# so each smart‑marker row is processed individually, binding a List<Employee> to markers and generating an Excel file with one row per employee.
// Keywords: Aspose.Cells | WorkbookDesigner | LineByLine | smart markers | C# example | list merging | Excel export | Aspose.Cells for .NET | populate rows from collection | template processing
// Common Searches: WorkbookDesigner LineByLine example C# | Aspose.Cells smart markers list merging | How to bind List<T> to smart markers Aspose.Cells | Generate Excel rows from collection using Aspose.Cells | LineByLine property usage Aspose.Cells .NET
// Developer Intent: Enable line‑by‑line processing to merge each item of a collection into separate rows using smart markers.
// Use Cases: Create an employee directory where each employee occupies its own row. | Produce a sales ledger that lists each transaction from a List<Sale>. | Generate invoices with line‑item details from a List<Product> while keeping header rows intact. | Export project task lists from a List<Task> into Excel for reporting.
// AI Prompts: Write C# code that adds a formatted header above the smart‑marker rows while keeping LineByLine true. | Show how to bind multiple collections (e.g., Employees and Departments) to the same worksheet using WorkbookDesigner with LineByLine enabled. | Compare the output of LineByLine true versus false for master‑detail smart markers and advise when to use each mode. | Provide a step‑by‑step guide to debug smart‑marker processing when LineByLine is set.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsLineByLineDemo
{
    // Simple data class for demonstration
    // Demonstrates enabling WorkbookDesigner.LineByLine in C# so each smart‑marker row is processed individually, binding a List<Employee> to markers and generating an Excel file with one row per employee.
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
            // 1. Create a new workbook and add smart markers for a simple list
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Smart markers placed line by line (each row will be repeated for each list item)
            sheet.Cells["A1"].PutValue("&Employees.Name");
            sheet.Cells["B1"].PutValue("&Employees.Age");
            sheet.Cells["C1"].PutValue("&Employees.Department");

            // 2. Prepare sample data source (a list of Employee objects)
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Age = 35, Department = "Sales" },
                new Employee { Name = "Jane Smith", Age = 28, Department = "Marketing" },
                new Employee { Name = "Bob Johnson", Age = 42, Department = "HR" }
            };

            // 3. Initialize WorkbookDesigner, assign the workbook, and enable line‑by‑line processing
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;
            designer.LineByLine = true; // Ensures each smart marker line is processed individually

            // 4. Bind the data source to the marker name used in the template
            designer.SetDataSource("Employees", employees);

            // 5. Process the smart markers to populate the worksheet
            designer.Process();

            // 6. Save the resulting workbook
            workbook.Save("Employees_LineByLine.xlsx", SaveFormat.Xlsx);
        }
    }
}
