using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomDataSourceDemo
{
    // Sample hierarchical classes
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

    // Custom ICellsDataTable implementation that flattens the hierarchy
    public class DeptEmployeeDataTable : ICellsDataTable
    {
        private readonly List<Department> _departments;
        private readonly List<(string DeptName, string EmpName, int EmpAge)> _rows;
        private int _currentRow = -1;

        public DeptEmployeeDataTable(List<Department> departments)
        {
            _departments = departments;
            _rows = new List<(string, string, int)>();

            // Flatten hierarchical data into rows
            foreach (var dept in _departments)
            {
                foreach (var emp in dept.Employees)
                {
                    _rows.Add((dept.DeptName, emp.Name, emp.Age));
                }
            }
        }

        // Indexer by row and column index
        public object this[int rowIndex, int columnIndex]
        {
            get
            {
                var row = _rows[rowIndex];
                return columnIndex switch
                {
                    0 => row.DeptName,
                    1 => row.EmpName,
                    2 => row.EmpAge,
                    _ => null
                };
            }
        }

        // Indexer by row (returns the whole row object)
        public object this[int rowIndex] => _rows[rowIndex];

        // Indexer by column name
        public object this[string columnName]
        {
            get
            {
                if (_currentRow < 0 || _currentRow >= _rows.Count) return null;
                var row = _rows[_currentRow];
                return columnName switch
                {
                    "DeptName" => row.DeptName,
                    "EmpName" => row.EmpName,
                    "EmpAge" => row.EmpAge,
                    _ => null
                };
            }
        }

        public int RowCount => _rows.Count;
        public int ColumnCount => 3;
        public int Count => _rows.Count;

        // Column names used in smart markers
        public string[] Columns => new[] { "DeptName", "EmpName", "EmpAge" };

        public void BeforeFirst()
        {
            _currentRow = -1;
        }

        public bool Next()
        {
            _currentRow++;
            return _currentRow < _rows.Count;
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
                sheet.Cells["A1"].PutValue("Department");
                sheet.Cells["B1"].PutValue("Employee");
                sheet.Cells["C1"].PutValue("Age");
                sheet.Cells["A2"].PutValue("&=DeptEmp.DeptName");
                sheet.Cells["B2"].PutValue("&=DeptEmp.EmpName");
                sheet.Cells["C2"].PutValue("&=DeptEmp.EmpAge");

                // Initialize WorkbookDesigner with the workbook
                var designer = new WorkbookDesigner(workbook);

                // Create custom data source object
                ICellsDataTable customTable = new DeptEmployeeDataTable(departments);

                // Bind the custom data source to a name used in smart markers
                designer.SetDataSource("DeptEmp", customTable);

                // Process the smart markers
                designer.Process();

                // Define output path and ensure directory exists
                string outputPath = "CustomHierarchicalDataOutput.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the result
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}