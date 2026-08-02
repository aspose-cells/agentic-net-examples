// Title: Custom hierarchical data source for WorkbookDesigner using ICellsDataTable (C# Aspose.Cells)
// Description: Demonstrates how to bind a complex department‑employee hierarchy to WorkbookDesigner by implementing a custom ICellsDataTable that flattens the nested collections. The example creates Department and Employee classes, builds a DepartmentDataSource, registers it with designer.SetDataSource("Dept", …), uses smart markers (&=Dept.DeptName, & =Dept.EmployeeName, & =Dept.EmployeeAge) and saves the populated workbook as an Excel file.
// Keywords: Aspose.Cells | WorkbookDesigner | custom data source | ICellsDataTable | C# | smart markers | hierarchical data | flatten nested collections | department employee report | Excel generation
// Common Searches: Aspose.Cells custom ICellsDataTable example | WorkbookDesigner hierarchical data source C# | bind nested objects to smart markers Aspose | flatten list of objects for smart markers | SetDataSource with custom data table Aspose
// Developer Intent: Bind a multi‑level object model (departments with employee lists) to WorkbookDesigner by providing a custom ICellsDataTable that presents the data in a flat tabular form for smart marker processing.
// Use Cases: Create an employee directory Excel report where each row shows department, employee name, and age via smart markers. | Reuse the same DepartmentDataSource across multiple worksheets or templates in a single reporting job. | Integrate the custom data source into a server‑side reporting service that converts nested business objects into Excel files.
// AI Prompts: Guide me through building a custom ICellsDataTable that flattens a list of departments and their employees for WorkbookDesigner. | Show how to extend DepartmentDataSource to add columns like EmployeeTitle and Salary while keeping smart marker compatibility. | What are common troubleshooting steps when smart markers return null or incorrect values from a custom data source?

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomDataSourceDemo
{
    // Hierarchical data model: Department contains a list of Employees
    // Demonstrates how to bind a complex department‑employee hierarchy to WorkbookDesigner by implementing a custom ICellsDataTable that flattens the nested collections. The example creates Department and Employee classes, builds a DepartmentDataSource, registers it with designer.SetDataSource("Dept", …), uses smart markers (&=Dept.DeptName, & =Dept.EmployeeName, & =Dept.EmployeeAge) and saves the populated workbook as an Excel file.
    public class Department
    {
        public string DeptName { get; set; }
        public List<Employee> Employees { get; set; }

        public Department(string name, List<Employee> employees)
        {
            DeptName = name;
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

    // Custom ICellsDataTable implementation that flattens the hierarchical structure
    public class DepartmentDataSource : ICellsDataTable
    {
        private readonly List<Department> _departments;
        private readonly List<(string DeptName, Employee Emp)> _flatRows;
        private int _currentRow = -1;

        public DepartmentDataSource(List<Department> departments)
        {
            _departments = departments ?? throw new ArgumentNullException(nameof(departments));
            _flatRows = new List<(string, Employee)>();

            // Flatten the hierarchy: each employee becomes a row with its department name
            foreach (var dept in _departments)
            {
                if (dept?.Employees == null) continue;
                foreach (var emp in dept.Employees)
                {
                    _flatRows.Add((dept.DeptName, emp));
                }
            }
        }

        // Indexer by row and column index
        public object this[int rowIndex, int columnIndex]
        {
            get
            {
                var row = _flatRows[rowIndex];
                return columnIndex == 0 ? (object)row.DeptName : row.Emp.Name;
            }
        }

        // Indexer by row (returns the whole row object, not used by Aspose but required)
        public object this[int rowIndex] => _flatRows[rowIndex];

        // Indexer by column name (used in smart markers)
        public object this[string columnName]
        {
            get
            {
                var row = _flatRows[_currentRow];
                return columnName switch
                {
                    "DeptName" => row.DeptName,
                    "EmployeeName" => row.Emp.Name,
                    "EmployeeAge" => row.Emp.Age,
                    _ => null
                };
            }
        }

        public int RowCount => _flatRows.Count;
        public int ColumnCount => 2; // DeptName, EmployeeName (Age accessed via column name)
        public int Count => _flatRows.Count;

        // Column names exposed to smart markers
        public string[] Columns => new[] { "DeptName", "EmployeeName", "EmployeeAge" };

        public void BeforeFirst()
        {
            _currentRow = -1;
        }

        public bool Next()
        {
            _currentRow++;
            return _currentRow < _flatRows.Count;
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Prepare hierarchical data
                var departments = new List<Department>
                {
                    new Department("HR", new List<Employee>
                    {
                        new Employee("Alice", 30),
                        new Employee("Bob", 28)
                    }),
                    new Department("IT", new List<Employee>
                    {
                        new Employee("Charlie", 35),
                        new Employee("Diana", 32)
                    })
                };

                // Create a new workbook and add smart markers
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Header row
                sheet.Cells["A1"].PutValue("Department");
                sheet.Cells["B1"].PutValue("Employee Name");
                sheet.Cells["C1"].PutValue("Employee Age");

                // Data rows with smart markers
                sheet.Cells["A2"].PutValue("&=Dept.DeptName");
                sheet.Cells["B2"].PutValue("&=Dept.EmployeeName");
                sheet.Cells["C2"].PutValue("&=Dept.EmployeeAge");

                // Initialize WorkbookDesigner with the workbook
                var designer = new WorkbookDesigner(workbook);

                // Set the custom hierarchical data source
                designer.SetDataSource("Dept", new DepartmentDataSource(departments));

                // Process smart markers
                designer.Process();

                // Define output file path
                string outputPath = "CustomHierarchicalDataSource.xlsx";

                // Save the result
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
