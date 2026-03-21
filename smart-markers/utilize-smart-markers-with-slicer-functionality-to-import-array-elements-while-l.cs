using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace SmartMarkerSlicerDemo
{
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.IgnoreUselessShapes = true;
            Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx", loadOptions);

            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Age = 30, Department = "Sales" },
                new Employee { Name = "Jane Smith", Age = 28, Department = "Marketing" },
                new Employee { Name = "Mike Johnson", Age = 35, Department = "Sales" },
                new Employee { Name = "Emily Davis", Age = 32, Department = "HR" }
            };

            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Employees", employees);
            designer.Process();

            Worksheet sheet = workbook.Worksheets[0];

            int firstRow = 0;
            int firstColumn = 0;
            int dataRows = employees.Count;
            int totalRows = dataRows + 1;
            int totalColumns = 3;

            int lastRow = firstRow + totalRows - 1;
            int lastColumn = firstColumn + totalColumns - 1;

            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "EmployeesTable";

            // Retrieve the "Department" column by index (0‑based). It is the third column.
            ListColumn deptColumn = table.ListColumns[2];

            sheet.Slicers.Add(table, deptColumn, 1, 4);

            workbook.Save("SmartMarkerSlicerResult.xlsx");
        }
    }
}