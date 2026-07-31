// Title: C# – Enable WorkbookDesigner.LineByLine for line‑by‑line smart‑marker list merging in Aspose.Cells
// Description: Demonstrates how to create an Excel template with smart markers, bind a List<Employee> to the "Employees" marker, activate WorkbookDesigner.LineByLine, process the markers so each employee occupies a separate row, and save the result while preserving any master markers.
// Keywords: Aspose.Cells | WorkbookDesigner | LineByLine | smart markers | C# list merging | Excel template population | SetDataSource | List<T> to Excel | .NET
// Common Searches: Aspose.Cells WorkbookDesigner LineByLine example | How to merge a list into Excel using smart markers C# | Preserve master markers when processing smart markers line by line | Smart marker list merging Aspose.Cells .NET | WorkbookDesigner SetDataSource List<T>
// Developer Intent: Activate WorkbookDesigner.LineByLine to insert each item of a simple list into consecutive rows while keeping existing master markers intact.
// Use Cases: Generate an employee roster where each employee appears on a new row using a List<Employee>. | Create invoices that add line‑items from a collection without overwriting the invoice header markers. | Build multi‑section reports that sequentially merge several data tables while retaining section titles.
// AI Prompts: Convert the List<Employee> data source to a DataTable while keeping LineByLine enabled. | Explain how WorkbookDesigner.LineByLine changes marker processing and how to switch back to the default mode. | Provide troubleshooting steps when LineByLine processing does not add rows as expected.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Demonstrates how to create an Excel template with smart markers, bind a List<Employee> to the "Employees" marker, activate WorkbookDesigner.LineByLine, process the markers so each employee occupies a separate row, and save the result while preserving any master markers.
    public class SimpleListMergeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook that will act as the template
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define smart markers for a simple list (one row per list item)
                sheet.Cells["A1"].PutValue("&=Employees.Name");
                sheet.Cells["B1"].PutValue("&=Employees.Age");
                sheet.Cells["C1"].PutValue("&=Employees.Department");

                // Sample data source: a list of employee objects
                List<Employee> employees = new List<Employee>
                {
                    new Employee { Name = "John Doe", Age = 30, Department = "Sales" },
                    new Employee { Name = "Jane Smith", Age = 28, Department = "HR" },
                    new Employee { Name = "Bob Johnson", Age = 35, Department = "IT" }
                };

                // Initialize the WorkbookDesigner, assign the workbook, and enable line‑by‑line processing
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook,
                    LineByLine = true // Obsolete, kept for compatibility
                };

                // Bind the data source to the smart marker name used in the template
                designer.SetDataSource("Employees", employees);

                // Process the smart markers to populate the worksheet
                designer.Process();

                // Save the populated workbook
                string outputPath = "SimpleListMerge_Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during processing: {ex.Message}");
            }
        }

        // POCO class representing an employee record
        public class Employee
        {
            public string? Name { get; set; }
            public int Age { get; set; }
            public string? Department { get; set; }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SimpleListMergeDemo.Run();
        }
    }
}
