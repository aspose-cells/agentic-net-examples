using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerMasterDetailDemo
{
    // Simple data classes
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class Department
    {
        public string Name { get; set; }
        public List<Employee> Employees { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (template) in memory
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Set up header row
            cells["A1"].PutValue("Department");
            cells["B1"].PutValue("Employee");
            cells["C1"].PutValue("Age");

            // 3. Add smart markers for master‑detail repeat
            // Master marker for Department.Name
            cells["A2"].PutValue("&=DeptData.Name");
            // Nested markers for each Employee inside the current Department
            cells["B2"].PutValue("&=DeptData.Employees.Name");
            cells["C2"].PutValue("&=DeptData.Employees.Age");

            // 4. Prepare sample data
            var departments = new List<Department>
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
                        new Employee { Name = "Sara", Age = 32 },
                        new Employee { Name = "Tom", Age = 29 }
                    }
                }
            };

            // 5. Initialize WorkbookDesigner and bind the data source
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;
            // Bind the list with the name "DeptData" used in the smart markers
            designer.SetDataSource("DeptData", departments);

            // 6. Process the smart markers – this will expand rows for each department
            // and nest employee rows inside each department block
            designer.Process();

            // 7. Save the resulting workbook
            designer.Workbook.Save("MasterDetailOutput.xlsx");
        }
    }
}