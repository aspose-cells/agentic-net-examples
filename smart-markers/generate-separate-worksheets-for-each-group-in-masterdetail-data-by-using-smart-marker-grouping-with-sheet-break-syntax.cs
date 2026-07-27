using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerSheetBreakDemo
{
    // Simple master‑detail classes
    public class Department
    {
        public string Name { get; set; }
        public List<Employee> Employees { get; set; }
    }

    public class Employee
    {
        public string Name { get; set; }
        public string Title { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // ---------- Create a workbook (lifecycle rule: create) ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Template";

            // ---------- Build the template with smart markers ----------
            // Header row
            sheet.Cells["A1"].PutValue("Department");
            sheet.Cells["B1"].PutValue("Employee");
            sheet.Cells["C1"].PutValue("Title");

            // Smart marker group start with sheet‑break syntax.
            // The marker "&=Departments.Start" tells Aspose.Cells to start a new sheet
            // for each distinct Department record.
            sheet.Cells["A2"].PutValue("&=Departments.Start");

            // Department name (master data)
            sheet.Cells["A3"].PutValue("&=Departments.Name");

            // Detail rows for Employees belonging to the current Department.
            // The group markers "&=Employees.Start" / "&=Employees.End" repeat the rows
            // for each employee in the master record.
            sheet.Cells["B3"].PutValue("&=Employees.Start");
            sheet.Cells["B3"].PutValue("&=Employees.Name");
            sheet.Cells["C3"].PutValue("&=Employees.Title");
            sheet.Cells["D3"].PutValue("&=Employees.End");

            // End of the master group
            sheet.Cells["A4"].PutValue("&=Departments.End");

            // ---------- Prepare master‑detail data ----------
            var departments = new List<Department>
            {
                new Department
                {
                    Name = "Sales",
                    Employees = new List<Employee>
                    {
                        new Employee { Name = "John Doe", Title = "Sales Manager" },
                        new Employee { Name = "Jane Smith", Title = "Sales Executive" }
                    }
                },
                new Department
                {
                    Name = "HR",
                    Employees = new List<Employee>
                    {
                        new Employee { Name = "Alice Brown", Title = "HR Manager" },
                        new Employee { Name = "Bob White", Title = "Recruiter" }
                    }
                }
            };

            // ---------- Set up the designer and bind data sources ----------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // LineByLine must be false because we are using a named range for smart markers.
                LineByLine = false
            };

            // Bind master and detail data sources.
            designer.SetDataSource("Departments", departments);
            // The detail source is automatically resolved from the master collection,
            // but we also bind it explicitly for clarity.
            designer.SetDataSource("Employees", departments);

            // ---------- Process the smart markers ----------
            designer.Process();

            // ---------- Save the result (lifecycle rule: save) ----------
            workbook.Save("DepartmentsBySheet.xlsx");
        }
    }
}