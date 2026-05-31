using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerNestedBindingDemo
{
    // Define the data model with a nested object hierarchy
    public class Employee
    {
        public string Name { get; set; }
        public Address Address { get; set; }
    }

    public class Address
    {
        public string Street { get; set; }
        public string City { get; set; }
        public string Country { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Insert smart markers that use dot notation to access nested properties
            //    &=$Employee.Name          -> Employee's Name
            //    &=$Employee.Address.Street -> Employee's Address.Street
            //    &=$Employee.Address.City   -> Employee's Address.City
            //    &=$Employee.Address.Country-> Employee's Address.Country
            sheet.Cells["A1"].PutValue("&=$Employee.Name");
            sheet.Cells["A2"].PutValue("&=$Employee.Address.Street");
            sheet.Cells["A3"].PutValue("&=$Employee.Address.City");
            sheet.Cells["A4"].PutValue("&=$Employee.Address.Country");

            // 3. Prepare sample data
            var employees = new List<Employee>
            {
                new Employee
                {
                    Name = "John Doe",
                    Address = new Address
                    {
                        Street = "123 Main St",
                        City = "New York",
                        Country = "USA"
                    }
                },
                new Employee
                {
                    Name = "Jane Smith",
                    Address = new Address
                    {
                        Street = "456 Oak Ave",
                        City = "London",
                        Country = "UK"
                    }
                }
            };

            // 4. Set up the WorkbookDesigner, assign the data source, and process the smart markers
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            // The table name "Employee" matches the prefix used in the smart markers
            designer.SetDataSource("Employee", employees);
            designer.Process();

            // 5. Save the result
            workbook.Save("NestedSmartMarkersOutput.xlsx");
        }
    }
}