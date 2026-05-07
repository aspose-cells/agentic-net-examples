using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerIfDemo
{
    // Simple data class used as a data source
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
            // Load the template workbook that contains smart markers.
            // The template should have a smart marker like:
            //   &=[Employees].IF([Age] > 30, Name)
            // This will render the employee's Name only if Age > 30.
            Workbook workbook = new Workbook("template.xlsx");

            // Prepare sample data.
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Age = 28, Department = "Sales" },
                new Employee { Name = "Jane Smith", Age = 35, Department = "HR" },
                new Employee { Name = "Bob Johnson", Age = 42, Department = "IT" }
            };

            // Initialize the WorkbookDesigner and bind the data source.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            // Bind the list to the smart marker name "Employees".
            designer.SetDataSource("Employees", employees);

            // Process the smart markers. The IF parameter will conditionally render data.
            designer.Process();

            // Save the result.
            workbook.Save("output.xlsx");
        }
    }
}