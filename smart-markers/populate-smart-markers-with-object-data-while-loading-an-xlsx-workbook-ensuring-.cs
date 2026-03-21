using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerDemo
{
    // Sample data class that will be bound to smart markers
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
            // Load the template workbook that contains smart markers like &Employee.Name, &Employee.Age, &Employee.Department
            Workbook workbook = new Workbook("Template.xlsx");

            // Prepare a collection of objects to be used as the data source
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Age = 30, Department = "Sales" },
                new Employee { Name = "Jane Smith", Age = 28, Department = "Marketing" },
                new Employee { Name = "Bob Johnson", Age = 35, Department = "IT" }
            };

            // Initialize the WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Bind the collection to a smart marker name (e.g., "Employee")
            designer.SetDataSource("Employee", employees);

            // Process the smart markers – this will populate the worksheet with the object data
            designer.Process();

            // Save the resulting workbook
            workbook.Save("Result.xlsx");
        }
    }
}