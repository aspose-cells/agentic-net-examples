// Title: Create a master smart marker for departments with nested employee detail markers using Aspose.Cells WorkbookDesigner in C#
// Common Searches: how to use Aspose.Cells WorkbookDesigner to create master‑detail smart markers in C# | nested smart markers example for department and employee data in Aspose.Cells | generate Excel report with department header and employee rows using smart markers | convert POCO department and employee lists to DataSet for Aspose.Cells smart markers | Aspose.Cells C# code for hierarchical smart markers with data relations
// Tags: master‑detail smart markers Aspose.Cells | WorkbookDesigner hierarchical data processing | department employee Excel generation with smart markers | POCO to DataSet conversion for Aspose.Cells | DataRelation usage in Excel smart marker templates

using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerDemo
{
    // Simple POCO classes for data source
    // This example demonstrates how to build an Excel workbook with a master smart marker for departments and nested employee detail markers using Aspose.Cells WorkbookDesigner in C#. It converts Department and Employee POCO collections into a DataSet with related tables, creates a DataRelation for hierarchical processing, applies the data source to the designer, processes the smart markers, and saves the output as SmartMarker_Nested_Departments.xlsx.
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Department
    {
        public string Name { get; set; } = string.Empty;
        public List<Employee> Employees { get; set; } = new List<Employee>();
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set up the smart marker template
                sheet.Cells[0, 0].PutValue("${Departments}");
                sheet.Cells[1, 0].PutValue("Department:");
                sheet.Cells[1, 1].PutValue("${Name}");
                sheet.Cells[2, 0].PutValue("Employee Name");
                sheet.Cells[2, 1].PutValue("Employee Age");
                sheet.Cells[3, 0].PutValue("${Employees.Name}");
                sheet.Cells[3, 1].PutValue("${Employees.Age}");

                // Prepare sample data
                List<Department> departments = new List<Department>
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

                // Convert POCO data to DataSet for WorkbookDesigner (required by this API version)
                DataSet dataSet = new DataSet();

                // Departments table
                DataTable deptTable = new DataTable("Departments");
                deptTable.Columns.Add("DepartmentId", typeof(int));
                deptTable.Columns.Add("Name", typeof(string));
                dataSet.Tables.Add(deptTable);

                // Employees table
                DataTable empTable = new DataTable("Employees");
                empTable.Columns.Add("DepartmentId", typeof(int));
                empTable.Columns.Add("Name", typeof(string));
                empTable.Columns.Add("Age", typeof(int));
                dataSet.Tables.Add(empTable);

                // Populate tables and maintain relation via DepartmentId
                int deptId = 1;
                foreach (var dept in departments)
                {
                    deptTable.Rows.Add(deptId, dept.Name);
                    foreach (var emp in dept.Employees)
                    {
                        empTable.Rows.Add(deptId, emp.Name, emp.Age);
                    }
                    deptId++;
                }

                // Create relation for nested smart markers
                DataRelation relation = new DataRelation(
                    "Dept_Employees",
                    deptTable.Columns["DepartmentId"],
                    empTable.Columns["DepartmentId"]);
                dataSet.Relations.Add(relation);

                // Process the smart markers using WorkbookDesigner
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dataSet);
                designer.Process();

                // Save the resulting workbook
                string outputFile = "SmartMarker_Nested_Departments.xlsx";
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
