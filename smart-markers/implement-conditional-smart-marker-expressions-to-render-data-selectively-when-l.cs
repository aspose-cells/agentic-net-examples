using System;
using System.Collections.Generic;
using Aspose.Cells;

class ConditionalSmartMarkerDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add column headers
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");

        // Insert smart markers.
        // The first cell uses a conditional smart marker:
        // It will display the employee's name only when Age > 30.
        // The second cell simply outputs the Age value.
        sheet.Cells["A2"].PutValue("&=IF($Age>30,$Name)");
        sheet.Cells["B2"].PutValue("&=$Age");

        // Define the range that contains the smart markers.
        // When LineByLine is false, the designer processes the range named "_CellsSmartMarkers".
        sheet.Cells.CreateRange("A2:B2").Name = "_CellsSmartMarkers";

        // Prepare a list of employee objects as the data source.
        List<Employee> employees = new List<Employee>
        {
            new Employee { Name = "John", Age = 28 },
            new Employee { Name = "Alice", Age = 35 },
            new Employee { Name = "Bob", Age = 42 }
        };

        // Set up the WorkbookDesigner.
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook,
            LineByLine = false // Use range smart markers instead of line‑by‑line processing.
        };

        // Bind the data source to the name "Employees".
        designer.SetDataSource("Employees", employees);

        // Process the smart markers and populate the worksheet.
        designer.Process();

        // Save the resulting workbook.
        workbook.Save("ConditionalSmartMarkerResult.xlsx");
    }

    // Simple POCO class representing an employee.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}