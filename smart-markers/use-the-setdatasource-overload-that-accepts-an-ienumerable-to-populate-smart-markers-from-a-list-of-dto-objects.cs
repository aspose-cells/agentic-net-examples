using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // DTO class representing an employee
    public class EmployeeDto
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }

        public EmployeeDto(string name, int age, string department)
        {
            Name = name;
            Age = age;
            Department = department;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook that will act as the template
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define header cells
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["C1"].PutValue("Department");

            // Insert smart markers that reference the "Employee" data source
            sheet.Cells["A2"].PutValue("&=$Employee.Name");
            sheet.Cells["B2"].PutValue("&=$Employee.Age");
            sheet.Cells["C2"].PutValue("&=$Employee.Department");

            // Prepare a list of DTO objects (IEnumerable) to be used as the data source
            List<EmployeeDto> employees = new List<EmployeeDto>
            {
                new EmployeeDto("John Doe", 30, "Sales"),
                new EmployeeDto("Jane Smith", 28, "Marketing"),
                new EmployeeDto("Mike Johnson", 35, "IT")
            };

            // Initialize the WorkbookDesigner with the template workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Bind the list to the smart marker variable "Employee" using the IEnumerable overload
            designer.SetDataSource("Employee", employees);

            // Process the smart markers and populate the worksheet with data
            designer.Process();

            // Save the populated workbook
            workbook.Save("EmployeesSmartMarkers.xlsx");
        }
    }
}