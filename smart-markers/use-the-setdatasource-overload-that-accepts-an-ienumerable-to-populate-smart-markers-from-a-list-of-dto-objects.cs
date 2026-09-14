// Title: Populate Excel smart markers from a List<EmployeeDto> using Aspose.Cells WorkbookDesigner.SetDataSource(IEnumerable) in C#
// AI Prompts: Write C# code that creates an Excel workbook, adds smart marker cells for employee fields, binds a List<EmployeeDto> to the "Employee" marker with WorkbookDesigner.SetDataSource, processes the markers, and saves the file. | Show how to use Aspose.Cells to expand rows automatically based on an IEnumerable collection of DTO objects and output the result as an .xlsx workbook.
// Common Searches: how to bind a List<T> to smart markers in Aspose.Cells C# | Aspose.Cells SetDataSource with IEnumerable example | populate Excel template using smart markers from a DTO collection | C# generate rows dynamically with smart markers Aspose.Cells | export employee data to XLSX using smart markers
// Tags: Aspose.Cells SetDataSource IEnumerable binding | C# smart markers Excel generation | dynamic rows from DTO collection | Excel export of employee list using Aspose.Cells | populate worksheet with object collection

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // DTO class representing an employee
    // // Demonstrates creating a workbook, defining smart marker cells for Name, Department, and Age, binding a List<EmployeeDto> to the "Employee" marker via WorkbookDesigner.SetDataSource, processing the markers to expand rows for each employee, and saving the result as EmployeesSmartMarkers.xlsx.
    public class EmployeeDto
    {
        public string Name { get; set; }
        public string Department { get; set; }
        public int Age { get; set; }

        public EmployeeDto(string name, string department, int age)
        {
            Name = name;
            Department = department;
            Age = age;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a list of DTO objects
            List<EmployeeDto> employees = new List<EmployeeDto>
            {
                new EmployeeDto("Alice", "HR", 30),
                new EmployeeDto("Bob", "IT", 28),
                new EmployeeDto("Charlie", "Finance", 35)
            };

            // Create a new workbook (template) and add smart markers
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Header row
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Department");
            sheet.Cells["C1"].PutValue("Age");

            // Smart marker rows – they will be expanded for each item in the list
            sheet.Cells["A2"].PutValue("&=$Employee.Name");
            sheet.Cells["B2"].PutValue("&=$Employee.Department");
            sheet.Cells["C2"].PutValue("&=$Employee.Age");

            // Initialize the designer with the workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Bind the IEnumerable (list) to the smart marker variable "Employee"
            designer.SetDataSource("Employee", employees);

            // Process the smart markers and populate data
            designer.Process();

            // Save the populated workbook
            workbook.Save("EmployeesSmartMarkers.xlsx");
        }
    }
}
