// Title: Filter employee data with LINQ and populate Excel using Aspose.Cells WorkbookDesigner smart markers (C#)
// AI Prompts: Generate C# code that uses LINQ to select employees older than a given age and passes the filtered collection to WorkbookDesigner for smart‑marker processing. | Show how to create a reusable method that accepts any IEnumerable<T>, applies a LINQ predicate, and sets it as the data source for Aspose.Cells WorkbookDesigner. | Provide an example that modifies the smart‑marker row to output only the Name and Department fields after the collection has been filtered.
// Common Searches: aspnet linq filter collection before using Aspose.Cells WorkbookDesigner smart markers | c# example of binding a filtered list to Aspose.Cells smart markers | how to use LINQ Where clause with Aspose.Cells WorkbookDesigner SetDataSource | export filtered employee records to Excel using Aspose.Cells smart markers | Aspose.Cells smart markers with LINQ filtered IEnumerable
// Tags: LINQ filtering with Aspose.Cells WorkbookDesigner | smart markers data binding filtered collection | C# export filtered employees to Excel using Aspose.Cells | WorkbookDesigner SetDataSource after LINQ Where | Excel generation with smart markers and LINQ

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerLinqDemo
{
    // Simple POCO representing an employee
    // Demonstrates creating a list of Employee objects, applying a LINQ Where clause to keep only employees older than 30, placing smart markers in a worksheet, passing the filtered list to WorkbookDesigner via SetDataSource, processing the markers to generate rows, and saving the result as FilteredEmployees.xlsx.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare source data (could come from any source, e.g., DB)
            // -----------------------------------------------------------------
            List<Employee> allEmployees = new List<Employee>
            {
                new Employee { Name = "John Doe",   Age = 28, Department = "Sales" },
                new Employee { Name = "Jane Smith", Age = 35, Department = "HR" },
                new Employee { Name = "Bob Brown",  Age = 42, Department = "IT" },
                new Employee { Name = "Alice White",Age = 31, Department = "Finance" }
            };

            // -----------------------------------------------------------------
            // 2. Filter the collection using LINQ (e.g., only employees older than 30)
            // -----------------------------------------------------------------
            List<Employee> filteredEmployees = allEmployees
                .Where(e => e.Age > 30)
                .ToList();

            // -----------------------------------------------------------------
            // 3. Create a workbook and place smart markers.
            //    Row 2 contains the markers; they will be repeated for each item.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet sheet = workbook.Worksheets[0];              // first worksheet

            // Header row
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["C1"].PutValue("Department");

            // Smart marker row – will be expanded for each Employee in the data source
            sheet.Cells["A2"].PutValue("&Employees.Name");
            sheet.Cells["B2"].PutValue("&Employees.Age");
            sheet.Cells["C2"].PutValue("&Employees.Department");

            // -----------------------------------------------------------------
            // 4. Bind the filtered collection to the WorkbookDesigner and process.
            // -----------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner();     // create designer
            designer.Workbook = workbook;                          // assign workbook
            designer.SetDataSource("Employees", filteredEmployees); // bind filtered list
            designer.Process();                                    // process smart markers

            // -----------------------------------------------------------------
            // 5. Save the result.
            // -----------------------------------------------------------------
            workbook.Save("FilteredEmployees.xlsx");
        }
    }
}
