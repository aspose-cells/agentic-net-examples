// Title: Binding a hierarchical C# object model to WorkbookDesigner smart markers using SetDataSource in Aspose.Cells for .NET
// Common Searches: C# set custom hierarchical data source for Aspose.Cells WorkbookDesigner smart markers | Aspose.Cells WorkbookDesigner populate Excel template with nested collections | How to use SetDataSource with complex object graph in Aspose.Cells .NET | Generate Excel report from company department employee hierarchy using Aspose.Cells | Load Excel template and bind multi‑level C# objects to smart markers Aspose.Cells
// Tags: WorkbookDesigner SetDataSource hierarchical object | Aspose.Cells smart markers nested collections | populate Excel template from C# object graph | custom data source for Excel smart markers | C# hierarchical data binding Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomDataSource
{
    // Hierarchical data model
    // The example defines Company, Department, and Employee classes to represent a hierarchical data model, creates sample data, loads or creates an Excel template with smart‑marker placeholders, assigns the hierarchy to WorkbookDesigner via SetDataSource("company", company), processes the template to fill the markers, and saves the populated workbook as Result.xlsx.
    public class Company
    {
        public string Name { get; set; } = string.Empty;
        public List<Department> Departments { get; set; } = new();
    }

    public class Department
    {
        public string DeptName { get; set; } = string.Empty;
        public List<Employee> Employees { get; set; } = new();
    }

    public class Employee
    {
        public string FullName { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public double Salary { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create sample hierarchical data
                var company = new Company
                {
                    Name = "Tech Solutions",
                    Departments = new List<Department>
                    {
                        new Department
                        {
                            DeptName = "Research",
                            Employees = new List<Employee>
                            {
                                new Employee { FullName = "Alice Johnson", Title = "Scientist", Salary = 95000 },
                                new Employee { FullName = "Bob Smith", Title = "Engineer", Salary = 85000 }
                            }
                        },
                        new Department
                        {
                            DeptName = "Sales",
                            Employees = new List<Employee>
                            {
                                new Employee { FullName = "Carol White", Title = "Sales Manager", Salary = 75000 },
                                new Employee { FullName = "David Brown", Title = "Sales Rep", Salary = 55000 }
                            }
                        }
                    }
                };

                const string templatePath = "Template.xlsx";
                Workbook workbook;

                // Load existing template or create a simple one if missing
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    workbook = new Workbook();
                    var ws = workbook.Worksheets[0];
                    // Simple markers for demonstration
                    ws.Cells["A1"].PutValue("&company.Name");
                    ws.Cells["A3"].PutValue("&company.Departments.DeptName");
                    ws.Cells["B3"].PutValue("&company.Departments.Employees.FullName");
                    ws.Cells["C3"].PutValue("&company.Departments.Employees.Title");
                    ws.Cells["D3"].PutValue("&company.Departments.Employees.Salary");
                }

                // Initialize WorkbookDesigner with the loaded workbook
                var designer = new WorkbookDesigner(workbook);

                // Set the custom hierarchical data source (named "company")
                designer.SetDataSource("company", company);

                // Process the template to populate data
                designer.Process();

                // Save the populated workbook
                workbook.Save("Result.xlsx");
                Console.WriteLine("Workbook generated successfully: Result.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
