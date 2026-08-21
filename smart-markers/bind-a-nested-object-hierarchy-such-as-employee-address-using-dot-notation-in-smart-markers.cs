// Title: C# – Aspose.Cells Smart Markers: Bind Nested Employee.Address via Dot‑Notation
// Description: A concise C# sample that creates a workbook, inserts smart markers referencing a List<Employee>, and uses dot‑notation (e.g., $Employees.Address.City) to bind nested address fields. WorkbookDesigner processes the markers, expands rows for each employee, and saves the result as EmployeesWithAddress.xlsx.
// Keywords: Aspose.Cells | Smart Markers | dot notation | nested objects | C# example | Employee address export | WorkbookDesigner | Excel data binding | hierarchical data | POCO to Excel | global
// Common Searches: aspocells smart markers nested object example | bind employee address to excel using dot notation | c# populate excel from hierarchical collection | how to use workbookdesigner with nested properties | excel report employee directory aspocells
// Developer Intent: Generate an Excel sheet that lists employees together with their street, city, and zip code by mapping nested Address properties through smart markers.
// Use Cases: Produce a printable employee directory that includes full mailing addresses. | Export HR data with embedded address details for bulk mail merges. | Create a dynamic Excel report for compliance audits that requires hierarchical employee information.
// AI Prompts: Add a smart‑marker column for phone numbers while keeping the existing dot‑notation for address fields. | Show how to bind a deeper hierarchy such as $Employees.Department.Name using the same technique. | Replace the List<Employee> source with a DataTable and demonstrate that dot‑notation still resolves nested columns.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsSmartMarkerNestedExample
{
    // Simple POCO classes representing an employee hierarchy
    // A concise C# sample that creates a workbook, inserts smart markers referencing a List<Employee>, and uses dot‑notation (e.g., $Employees.Address.City) to bind nested address fields. WorkbookDesigner processes the markers, expands rows for each employee, and saves the result as EmployeesWithAddress.xlsx.
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public Address Address { get; set; } = new Address();
    }

    public class Address
    {
        public string Street { get; set; } = string.Empty;
        public string City   { get; set; } = string.Empty;
        public string Zip    { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Insert smart markers using dot notation for nested objects
                // Header row
                sheet.Cells["A1"].PutValue("Employee Name");
                sheet.Cells["B1"].PutValue("Street");
                sheet.Cells["C1"].PutValue("City");
                sheet.Cells["D1"].PutValue("Zip");

                // Data rows – use smart markers that reference the collection "Employees"
                sheet.Cells["A2"].PutValue("&=$Employees.Name");
                sheet.Cells["B2"].PutValue("&=$Employees.Address.Street");
                sheet.Cells["C2"].PutValue("&=$Employees.Address.City");
                sheet.Cells["D2"].PutValue("&=$Employees.Address.Zip");

                // Mark the range that contains smart markers (required for processing)
                AsposeRange smartRange = sheet.Cells.CreateRange("A2:D2");
                smartRange.Name = "_CellsSmartMarkers";

                // 3. Prepare sample data – a list of employees with nested address objects
                List<Employee> employees = new List<Employee>
                {
                    new Employee
                    {
                        Name = "John Doe",
                        Address = new Address
                        {
                            Street = "123 Main St",
                            City   = "New York",
                            Zip    = "10001"
                        }
                    },
                    new Employee
                    {
                        Name = "Jane Smith",
                        Address = new Address
                        {
                            Street = "456 Oak Ave",
                            City   = "Los Angeles",
                            Zip    = "90001"
                        }
                    },
                    new Employee
                    {
                        Name = "Bob Johnson",
                        Address = new Address
                        {
                            Street = "789 Pine Rd",
                            City   = "Chicago",
                            Zip    = "60601"
                        }
                    }
                };

                // 4. Create a WorkbookDesigner, assign the workbook, and bind the data source
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                // The name "Employees" matches the smart marker prefix $Employees
                designer.SetDataSource("Employees", employees);

                // 5. Process the smart markers – this will expand the range for each employee
                designer.Process();

                // 6. Save the populated workbook
                string outputPath = "EmployeesWithAddress.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
