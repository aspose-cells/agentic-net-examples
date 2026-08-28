// Title: Use dot‑notation smart markers to bind a nested Employee‑Address object hierarchy in Aspose.Cells for .NET
// AI Prompts: Create a new Workbook, insert smart markers '&=$Employee.Name', '&=$Employee.Age', '&=$Employee.Address.Street', '&=$Employee.Address.City', '&=$Employee.Address.Zip', bind an Employee instance (including its Address) with WorkbookDesigner.SetDataSource, call WorkbookDesigner.Process, and save the workbook as an .xlsx file. | Show how to populate an Excel template from a hierarchical object (Employee containing Address) by adding dot‑notation smart markers, assigning the object as the data source, processing the markers with WorkbookDesigner, and exporting the result.
// Common Searches: asp.net how to bind nested object to Excel using Aspose.Cells smart markers | dot notation syntax for accessing child properties in Aspose.Cells smart markers | populate employee address fields in Excel with Aspose.Cells WorkbookDesigner C# | example of using &=$Employee.Address.City smart marker in Aspose.Cells | generate Excel file from hierarchical POCO with Aspose.Cells smart markers
// Tags: Aspose.Cells smart markers hierarchical syntax | bind nested object hierarchy to Excel | employee address smart marker example | C# populate hierarchical data in Excel | process smart markers with nested objects

using System;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerNestedExample
{
    // Simple POCO classes representing a nested object hierarchy
    // The sample creates a workbook, places smart markers that reference Employee.Name, Employee.Age, and the nested Address properties using dot notation, binds an Employee object (with an Address) as the data source via WorkbookDesigner, processes the markers, and saves the populated spreadsheet as EmployeeNestedSmartMarkers.xlsx.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public Address Address { get; set; }
    }

    public class Address
    {
        public string Street { get; set; }
        public string City { get; set; }
        public string Zip { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook and add smart markers using dot notation
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();                     // create a new workbook
            Worksheet sheet = workbook.Worksheets[0];               // get the first worksheet

            // Place smart markers in cells. The dot notation accesses nested properties.
            sheet.Cells["A1"].PutValue("&=$Employee.Name");          // employee name
            sheet.Cells["A2"].PutValue("&=$Employee.Age");           // employee age
            sheet.Cells["A3"].PutValue("&=$Employee.Address.Street"); // street
            sheet.Cells["A4"].PutValue("&=$Employee.Address.City");   // city
            sheet.Cells["A5"].PutValue("&=$Employee.Address.Zip");    // zip code

            // -----------------------------------------------------------------
            // 2. Prepare the nested data source
            // -----------------------------------------------------------------
            Employee emp = new Employee
            {
                Name = "John Doe",
                Age = 35,
                Address = new Address
                {
                    Street = "123 Main St",
                    City = "Metropolis",
                    Zip = "12345"
                }
            };

            // -----------------------------------------------------------------
            // 3. Bind the data source to the workbook designer and process markers
            // -----------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Bind the root object with the name "Employee"
            designer.SetDataSource("Employee", emp);

            // Process all smart markers in the workbook
            designer.Process();

            // -----------------------------------------------------------------
            // 4. Save the result
            // -----------------------------------------------------------------
            workbook.Save("EmployeeNestedSmartMarkers.xlsx");
        }
    }
}
