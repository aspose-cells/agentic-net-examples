// Title: Create an Excel report from nested Department and Employee objects using Aspose.Cells smart marker range syntax in C#
// AI Prompts: Write a C# program that uses Aspose.Cells WorkbookDesigner to populate a worksheet with Department names and their Employees' names and ages via smart markers. | Modify the smart marker range to add extra employee columns such as Position and Salary, and adjust the data source mapping accordingly. | Add logic to accept a user‑specified output path, create the target folder if needed, and save the generated Excel file with Aspose.Cells.
// Common Searches: aspnet cells smart markers nested collection example c# | how to use smart marker range for parent child objects in Aspose.Cells | c# generate excel from hierarchical data using WorkbookDesigner smart markers | aspose.cells map department employee list to excel rows | c# smart marker syntax for nested lists Aspose.Cells
// Tags: smart markers nested collections Aspose.Cells | WorkbookDesigner populate Excel from hierarchical objects | C# generate Excel using smart marker range | auto expand rows Aspose.Cells smart markers | Excel report department employee hierarchy

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerDemo
{
    // Define the nested object hierarchy
    // The example defines Department and Employee classes, builds sample data, creates a workbook, sets header cells, inserts smart markers (&=Departments.Name, &=Departments.Employees.Name, &=Departments.Employees.Age) in a single row range, assigns the department list as the data source to a WorkbookDesigner, processes the markers to auto‑expand rows for each employee, ensures the output directory exists, and saves the file as NestedSmartMarkerReport.xlsx.
    public class Department
    {
        public string Name { get; set; } = string.Empty;
        public List<Employee> Employees { get; set; } = new List<Employee>();
    }

    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Prepare sample data
                var departments = new List<Department>
                {
                    new Department
                    {
                        Name = "Sales",
                        Employees = new List<Employee>
                        {
                            new Employee { Name = "Alice", Age = 30 },
                            new Employee { Name = "Bob", Age = 28 }
                        }
                    },
                    new Department
                    {
                        Name = "Engineering",
                        Employees = new List<Employee>
                        {
                            new Employee { Name = "Charlie", Age = 35 },
                            new Employee { Name = "Diana", Age = 32 },
                            new Employee { Name = "Ethan", Age = 29 }
                        }
                    }
                };

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Report";

                // Set header titles
                sheet.Cells["A1"].PutValue("Department");
                sheet.Cells["B1"].PutValue("Employee");
                sheet.Cells["C1"].PutValue("Age");

                // Insert smart markers.
                // Row 2 will be the start of the smart marker range.
                // The marker syntax "&=Collection.Property" tells Aspose.Cells to repeat the row
                // for each item in the collection. Nested collections are handled by using
                // the parent collection name followed by the child collection name.
                sheet.Cells["A2"].PutValue("&=Departments.Name");
                sheet.Cells["B2"].PutValue("&=Departments.Employees.Name");
                sheet.Cells["C2"].PutValue("&=Departments.Employees.Age");

                // Define the smart marker range (from row 2 to row 2)
                // This range will be expanded automatically based on the data source.
                AsposeRange smartRange = sheet.Cells.CreateRange("A2:C2");
                // smartRange.IsSmartMarker = true; // Not required for processing

                // Set the data source for the smart markers using WorkbookDesigner.
                // The key "Departments" must match the collection name used in the markers.
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Departments", departments);

                // Process the smart markers – this will populate the worksheet with the data.
                designer.Process();

                // Save the workbook to a file
                string outputPath = "NestedSmartMarkerReport.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Report generated successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
