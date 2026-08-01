// Title: C# – Generate Excel with nested Department/Employee data using Aspose.Cells range smart markers
// Description: Shows how to create a workbook, define a range (A2:C2) with smart markers, bind a List<Department> (each containing Employees) to WorkbookDesigner, process the range markers to auto‑expand rows for every employee, and save the result.
// Keywords: Aspose.Cells | C# smart markers | range smart markers | nested collections | WorkbookDesigner | Excel export .NET | hierarchical data | Department Employee example | GitHub code sample
// Common Searches: Aspose.Cells range smart markers nested collection | C# generate Excel from hierarchical objects | WorkbookDesigner parent child smart markers example | How to export department employee list to Excel using Aspose.Cells | Range smart marker syntax for child collections
// Developer Intent: Generate an Excel workbook that automatically expands rows for each employee under their department using Aspose.Cells range smart markers in C#.
// Use Cases: Departmental staff directory report | Payroll sheet grouped by department | Organizational chart export to Excel | Sales and HR summary with hierarchical grouping | Any .NET application needing hierarchical Excel export without manual loops
// AI Prompts: Write C# code using Aspose.Cells WorkbookDesigner to process range smart markers for a parent collection 'Category' with a child collection 'Products'. | Show how to add a 'Position' column to the Department/Employee smart marker example while keeping automatic row expansion. | Generate NUnit unit tests that verify the row count for each department after processing nested smart markers. | Explain the difference between cell smart markers and range smart markers for nested data in Aspose.Cells. | Provide a step‑by‑step guide to publish this example on GitHub with a proper README and SEO‑friendly metadata.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using CellsRange = Aspose.Cells.Range;

namespace AsposeCellsNestedSmartMarkersDemo
{
    // Sample data classes
    // Shows how to create a workbook, define a range (A2:C2) with smart markers, bind a List<Department> (each containing Employees) to WorkbookDesigner, process the range markers to auto‑expand rows for every employee, and save the result.
    public class Department
    {
        public string Name { get; set; }
        public List<Employee> Employees { get; set; }

        public Department(string name, List<Employee> employees)
        {
            Name = name;
            Employees = employees;
        }
    }

    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }

        public Employee(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // 3. Set up smart markers using range smart marker syntax
                // Header row
                sheet.Cells["A1"].PutValue("Department");
                sheet.Cells["B1"].PutValue("Employee Name");
                sheet.Cells["C1"].PutValue("Employee Age");

                // Data rows (the range that will be processed)
                // Row 2 – parent marker
                sheet.Cells["A2"].PutValue("&Department.Name");
                // Row 2 – child markers (will be repeated for each employee)
                sheet.Cells["B2"].PutValue("&Department.Employees.Name");
                sheet.Cells["C2"].PutValue("&Department.Employees.Age");

                // Define the range that contains the smart markers and name it "_CellsSmartMarkers"
                CellsRange smartRange = sheet.Cells.CreateRange("A2:C2");
                smartRange.Name = "_CellsSmartMarkers";

                // 4. Prepare nested data
                var departments = new List<Department>
                {
                    new Department(
                        "Sales",
                        new List<Employee>
                        {
                            new Employee("John", 30),
                            new Employee("Emma", 28)
                        }),
                    new Department(
                        "HR",
                        new List<Employee>
                        {
                            new Employee("Mike", 35),
                            new Employee("Sara", 32)
                        })
                };

                // 5. Initialize WorkbookDesigner and assign the workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // 6. Set the data source for the parent collection.
                // The name "Department" matches the smart marker prefix.
                designer.SetDataSource("Department", departments);

                // 7. Process the smart markers (range smart markers are used, so no parameters needed)
                designer.Process();

                // 8. Save the result
                string outputPath = "NestedSmartMarkersOutput.xlsx";
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
