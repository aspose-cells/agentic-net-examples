using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerSheetBreakDemo
{
    // Sample master‑detail classes
    public class Department
    {
        public string Name { get; set; }
        public List<Employee> Employees { get; set; }
    }

    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // 2. Get the first worksheet to design the template
                Worksheet sheet = workbook.Worksheets[0];

                // 3. Build the template using smart markers and sheet‑break syntax
                // Row 1 – headers for master data (Department)
                sheet.Cells["A1"].PutValue("Department");
                // Row 2 – sheet break marker for each Department group
                // Correct syntax: &=[DataSourceName] creates a new sheet per group
                sheet.Cells["A2"].PutValue("&=[Departments]");
                // Row 3 – master field (Department name)
                sheet.Cells["A3"].PutValue("&=Name");

                // Row 5 – headers for detail data (Employees)
                sheet.Cells["A5"].PutValue("Employee Name");
                sheet.Cells["B5"].PutValue("Age");
                // Row 6 – detail fields; hierarchical path to the child collection
                sheet.Cells["A6"].PutValue("&=Employees.Name");
                sheet.Cells["B6"].PutValue("&=Employees.Age");

                // 4. Prepare sample master‑detail data
                List<Department> departments = new List<Department>
                {
                    new Department
                    {
                        Name = "Sales",
                        Employees = new List<Employee>
                        {
                            new Employee { Name = "John", Age = 30 },
                            new Employee { Name = "Emma", Age = 28 }
                        }
                    },
                    new Department
                    {
                        Name = "HR",
                        Employees = new List<Employee>
                        {
                            new Employee { Name = "Mike", Age = 35 },
                            new Employee { Name = "Sara", Age = 32 }
                        }
                    }
                };

                // 5. Set the data source for the smart markers
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                // Bind the list to the name used in the sheet‑break marker
                designer.SetDataSource("Departments", departments);

                // 6. Process the smart markers (lifecycle rule: process)
                designer.Process();

                // 7. Save the result (lifecycle rule: save)
                workbook.Save("DepartmentsBySheet.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}