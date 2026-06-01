using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsLineByLineDemo
{
    // Simple data class for demonstration
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
            // Load the template workbook that contains smart markers
            Workbook workbook = new Workbook("template.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Ensure processing is done line‑by‑line (default is true, set explicitly for clarity)
            designer.LineByLine = true;

            // Prepare sample data source (a simple list)
            var employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Age = 35, Department = "Sales" },
                new Employee { Name = "Jane Smith", Age = 28, Department = "Marketing" }
            };

            // Bind the data source to a smart marker name used in the template
            designer.SetDataSource("Employees", employees);

            // Process the smart markers and populate the worksheet
            designer.Process();

            // Save the resulting workbook
            workbook.Save("output.xlsx");
        }
    }
}