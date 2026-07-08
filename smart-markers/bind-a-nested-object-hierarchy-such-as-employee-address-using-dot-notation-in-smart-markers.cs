using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerNestedDemo
{
    // Simple POCO classes representing an employee hierarchy
    public class Employee
    {
        public string Name { get; set; }
        public Address Address { get; set; }
    }

    public class Address
    {
        public string Street { get; set; }
        public string City   { get; set; }
        public string State  { get; set; }
        public string Zip    { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook (template)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Insert smart markers using dot notation for nested objects
            // Header row
            sheet.Cells["A1"].PutValue("Employee Name");
            sheet.Cells["B1"].PutValue("Street");
            sheet.Cells["C1"].PutValue("City");
            sheet.Cells["D1"].PutValue("State");
            sheet.Cells["E1"].PutValue("Zip");

            // Data row – the range will be processed as a repeating block
            sheet.Cells["A2"].PutValue("&=$Employees.Name");
            sheet.Cells["B2"].PutValue("&=$Employees.Address.Street");
            sheet.Cells["C2"].PutValue("&=$Employees.Address.City");
            sheet.Cells["D2"].PutValue("&=$Employees.Address.State");
            sheet.Cells["E2"].PutValue("&=$Employees.Address.Zip");

            // Mark the range that contains smart markers
            // The name "_CellsSmartMarkers" tells Aspose.Cells to treat the range as a smart‑marker block
            sheet.Cells.CreateRange("A2:E2").Name = "_CellsSmartMarkers";

            // 3. Prepare sample data – a list of employees with nested address objects
            List<Employee> employees = new List<Employee>
            {
                new Employee
                {
                    Name = "John Doe",
                    Address = new Address
                    {
                        Street = "123 Main St",
                        City   = "Springfield",
                        State  = "IL",
                        Zip    = "62701"
                    }
                },
                new Employee
                {
                    Name = "Jane Smith",
                    Address = new Address
                    {
                        Street = "456 Oak Ave",
                        City   = "Metropolis",
                        State  = "NY",
                        Zip    = "10001"
                    }
                }
            };

            // 4. Set up the WorkbookDesigner and bind the data source
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            // Bind the list to the name used in the smart markers ("Employees")
            designer.SetDataSource("Employees", employees);

            // 5. Process the smart markers – this expands the marked range for each employee
            designer.Process();

            // 6. Save the populated workbook
            workbook.Save("NestedSmartMarkersOutput.xlsx");
        }
    }
}