using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (template)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Merge cells A1:C1 for a header title
        cells.Merge(0, 0, 1, 3);
        cells[0, 0].PutValue("Employee Report");

        // Define a merged template row (A2:C2) that will be duplicated for each data item
        cells.Merge(1, 0, 1, 3);
        // Place a smart marker inside the merged cell; it will be expanded per data row
        cells[1, 0].PutValue("&=Employees.Name");

        // Additional columns (not merged) for other fields
        cells[1, 3].PutValue("&=Employees.Age");
        cells[1, 4].PutValue("&=Employees.Department");

        // Sample data source
        List<Employee> data = new List<Employee>
        {
            new Employee { Name = "John Doe", Age = 30, Department = "Sales" },
            new Employee { Name = "Jane Smith", Age = 28, Department = "HR" },
            new Employee { Name = "Bob Johnson", Age = 35, Department = "IT" }
        };

        // Set up WorkbookDesigner, assign the data source and process smart markers
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;
        designer.SetDataSource("Employees", data);
        // Process will expand the merged row for each employee automatically
        designer.Process();

        // Save the populated workbook
        workbook.Save("MergedSmartMarkersOutput.xlsx");
    }

    // Simple POCO for employee data
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }
    }
}