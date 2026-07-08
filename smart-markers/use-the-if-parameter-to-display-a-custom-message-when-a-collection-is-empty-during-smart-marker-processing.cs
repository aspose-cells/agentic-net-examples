using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsIfSmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Smart marker with IF parameter:
            // If the collection "Employees" is empty, display a custom message.
            // Otherwise, the message is an empty string and the data rows will be populated.
            cells["A1"].PutValue("&IF=Employees?\"No employee data available\":\"\"");

            // Data rows (will be filled only when the collection has items)
            cells["A2"].PutValue("&=$Employees.Name");
            cells["B2"].PutValue("&=$Employees.Age");

            // Create an empty data source (no employees)
            List<Employee> employees = new List<Employee>(); // empty collection

            // Set up the WorkbookDesigner with the workbook and data source
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = wb;
            designer.SetDataSource("Employees", employees);

            // Process the smart markers
            designer.Process();

            // Save the resulting workbook
            wb.Save("IfSmartMarkerOutput.xlsx");
        }
    }

    // Simple POCO class representing an employee
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}