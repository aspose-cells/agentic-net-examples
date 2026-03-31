using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Nested data classes
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public Address Address { get; set; }
    }

    public class Address
    {
        public string City { get; set; }
        public string ZipCode { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // Step 1: Create a template workbook with smart markers that use
            // dot notation to reference nested object properties.
            // -----------------------------------------------------------------
            Workbook template = new Workbook();
            Worksheet sheet = template.Worksheets[0];
            Cells cells = sheet.Cells;

            // Smart markers using dot notation
            cells["A1"].PutValue("&=Employee.Name");
            cells["B1"].PutValue("&=Employee.Age");
            cells["A2"].PutValue("&=Employee.Address.City");
            cells["B2"].PutValue("&=Employee.Address.ZipCode");

            // Save the template to a temporary file
            string templatePath = Path.Combine(Path.GetTempPath(), "EmployeeTemplate.xlsx");
            template.Save(templatePath, SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // Step 2: Load the template workbook using WorkbookDesigner.
            // -----------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = new Workbook(templatePath);

            // -----------------------------------------------------------------
            // Step 3: Prepare data source with nested objects.
            // -----------------------------------------------------------------
            var employees = new List<Employee>
            {
                new Employee
                {
                    Name = "John Doe",
                    Age = 30,
                    Address = new Address { City = "New York", ZipCode = "10001" }
                },
                new Employee
                {
                    Name = "Jane Smith",
                    Age = 28,
                    Address = new Address { City = "Los Angeles", ZipCode = "90001" }
                }
            };

            // Bind the data source. The table name "Employee" matches the smart marker prefix.
            designer.SetDataSource("Employee", employees);

            // -----------------------------------------------------------------
            // Step 4: Process the smart markers. The dot notation will resolve
            // nested properties automatically.
            // -----------------------------------------------------------------
            designer.Process();

            // -----------------------------------------------------------------
            // Step 5: Save the populated workbook.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(Path.GetTempPath(), "EmployeeReport.xlsx");
            designer.Workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Template created at: {templatePath}");
            Console.WriteLine($"Report generated at: {outputPath}");
        }
    }
}