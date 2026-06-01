using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsSmartMarkerNested
{
    // Data model classes
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Department
    {
        public string DeptName { get; set; } = string.Empty;
        public List<Employee> Employees { get; set; } = new();
    }

    public class Company
    {
        public string Name { get; set; } = string.Empty;
        public List<Department> Departments { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create sample nested data
                var company = new Company
                {
                    Name = "Tech Corp",
                    Departments = new List<Department>
                    {
                        new Department
                        {
                            DeptName = "Research",
                            Employees = new List<Employee>
                            {
                                new Employee { Name = "Alice", Age = 30 },
                                new Employee { Name = "Bob", Age = 28 }
                            }
                        },
                        new Department
                        {
                            DeptName = "Development",
                            Employees = new List<Employee>
                            {
                                new Employee { Name = "Charlie", Age = 35 },
                                new Employee { Name = "Diana", Age = 32 }
                            }
                        }
                    }
                };

                // Create a new workbook (no template file needed)
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Header row for company name
                sheet.Cells["A1"].PutValue("Company");
                sheet.Cells["B1"].PutValue("&=Company.Name");

                // Row for department name (repeated for each department)
                sheet.Cells["A3"].PutValue("Department");
                sheet.Cells["B3"].PutValue("&=Company.Departments.DeptName");

                // Row for employee details (repeated for each employee within a department)
                sheet.Cells["A4"].PutValue("Employee");
                sheet.Cells["B4"].PutValue("&=Company.Departments.Employees.Name");
                sheet.Cells["C4"].PutValue("&=Company.Departments.Employees.Age");

                // Define the range that contains the repeating rows and name it for range smart markers
                AsposeRange smartRange = sheet.Cells.CreateRange("A3:C4");
                smartRange.Name = "_CellsSmartMarkers";

                // Initialize WorkbookDesigner and assign the workbook
                var designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Set data sources
                designer.SetDataSource("Company", company);
                designer.SetDataSource("Departments", company.Departments);
                // Employees collection is accessed via nested markers; explicit binding is optional

                // Process the smart markers
                designer.Process();

                // Ensure output directory exists
                string outputPath = "NestedSmartMarkersOutput.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the populated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}