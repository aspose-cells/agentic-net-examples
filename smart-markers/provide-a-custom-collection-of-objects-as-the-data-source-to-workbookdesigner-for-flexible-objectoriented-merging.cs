// Title: How to bind a List<Employee> POCO collection to WorkbookDesigner for smart‑marker merging in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set data source for smart markers using a List of custom objects | Merge a collection of Employee objects into an Excel template with WorkbookDesigner | Load an Excel template and populate smart markers from a POCO list in .NET | WorkbookDesigner SetDataSource example with a custom class collection | Smart marker merging multiple rows from List<Employee> in Aspose.Cells
// Tags: WorkbookDesigner SetDataSource C# | smart markers custom object collection | Excel template merge POCO list Aspose.Cells | populate smart markers from List<Employee> | Aspose.Cells data binding for Excel templates

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomDataSource
{
    // Define a custom class that will serve as the data source.
    // The example defines an Employee POCO class, loads an Excel template that contains smart‑marker fields, creates a List<Employee> with sample data, assigns the list to the WorkbookDesigner via SetDataSource using the "Employees" prefix, processes the smart markers to fill the worksheet, and saves the merged workbook as Result.xlsx.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public double Salary { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new Workbook instance (or load an existing template).
            //    The template should contain merge fields like &=Employees.Name, &=Employees.Age, etc.
            Workbook workbook = new Workbook("Template.xlsx"); // Load template

            // 2. Prepare a collection of custom objects.
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "Alice", Age = 30, Salary = 75000 },
                new Employee { Name = "Bob",   Age = 45, Salary = 88000 },
                new Employee { Name = "Carol", Age = 28, Salary = 62000 }
            };

            // 3. Create a WorkbookDesigner and set the data source.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Employees", employees); // "Employees" matches the merge field prefix in the template

            // 4. Process the smart markers to merge data into the worksheet.
            designer.Process();

            // 5. Save the resulting workbook.
            workbook.Save("Result.xlsx");
        }
    }
}
