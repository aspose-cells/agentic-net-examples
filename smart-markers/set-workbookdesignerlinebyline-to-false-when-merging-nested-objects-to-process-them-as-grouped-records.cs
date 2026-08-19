// Title: Set WorkbookDesigner.LineByLine = false to group nested objects with smart markers (C#)
// Description: Demonstrates how to disable line‑by‑line processing in Aspose.Cells WorkbookDesigner so that a root collection and its nested Department list are merged as grouped records in an Excel workbook using smart markers.
// Keywords: Aspose.Cells WorkbookDesigner LineByLine false | C# smart markers nested collection | grouped records Excel export | hierarchical data Aspose.Cells | merge parent child records | WorkbookDesigner hierarchical data | smart markers C# example
// Common Searches: Aspose.Cells disable line by line processing | WorkbookDesigner group nested objects | smart markers hierarchical data C# | LineByLine false example Aspose.Cells | export master‑detail Excel with Aspose.Cells
// Developer Intent: Turn off line‑by‑line mode so that child collections are rendered as grouped rows under each parent record.
// Use Cases: Create a master‑detail Excel report where each employee appears once with their departments listed below. | Export orders with line items while preserving the order‑item hierarchy. | Generate a consolidated financial sheet that groups department budgets under each company.
// AI Prompts: Show how to set WorkbookDesigner.LineByLine to false and bind a list with nested collections using smart markers in C#. | Explain the difference between LineByLine true and false in WorkbookDesigner and illustrate the grouped output. | Provide a complete C# example that exports hierarchical data to Excel with WorkbookDesigner while disabling line‑by‑line processing.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsLineByLineDemo
{
    // Demonstrates how to disable line‑by‑line processing in Aspose.Cells WorkbookDesigner so that a root collection and its nested Department list are merged as grouped records in an Excel workbook using smart markers.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define smart markers for a root collection and its nested collection
            // Root level markers
            sheet.Cells["A1"].PutValue("&RootData.Name");
            sheet.Cells["B1"].PutValue("&RootData.Age");
            // Nested collection markers (departments)
            sheet.Cells["A2"].PutValue("&RootData.Departments.DName");
            sheet.Cells["B2"].PutValue("&RootData.Departments.Budget");

            // Sample data with nested objects
            var data = new List<RootData>
            {
                new RootData
                {
                    Name = "John",
                    Age = 30,
                    Departments = new List<Department>
                    {
                        new Department { DName = "Sales", Budget = 100000 },
                        new Department { DName = "HR", Budget = 50000 }
                    }
                },
                new RootData
                {
                    Name = "Alice",
                    Age = 28,
                    Departments = new List<Department>
                    {
                        new Department { DName = "IT", Budget = 150000 }
                    }
                }
            };

            // Initialize WorkbookDesigner, assign the workbook, and set LineByLine to false
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                LineByLine = false // Process nested objects as grouped records
            };

            // Bind the root collection to the smart marker name
            designer.SetDataSource("RootData", data);

            // Process the smart markers
            designer.Process();

            // Save the processed workbook
            workbook.Save("LineByLineFalseOutput.xlsx");
        }
    }

    // Root data class containing a collection of nested objects
    public class RootData
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public List<Department> Departments { get; set; }
    }

    // Nested object class representing a department
    public class Department
    {
        public string DName { get; set; }
        public double Budget { get; set; }
    }
}
