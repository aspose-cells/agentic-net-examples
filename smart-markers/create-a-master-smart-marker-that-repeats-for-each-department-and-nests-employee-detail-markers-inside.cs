// Title: Generate a nested master‑detail Excel report with department and employee smart markers using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example for nested smart markers with master‑detail relationship | How to use WorkbookDesigner to create department and employee hierarchy in Excel | Aspose.Cells smart markers repeat for each department and its employees C# | Generating Excel report from DataSet relations using smart markers in Aspose.Cells | Set up DataSet relation for nested smart markers in Aspose.Cells C#
// Tags: nested smart markers with WorkbookDesigner | master‑detail Excel generation Aspose.Cells | populate Excel from DataSet relations C# | department employee hierarchy smart markers | repeat rows using smart markers Aspose.Cells | C# Aspose.Cells nested markers example

using System;
using System.Data;
using Aspose.Cells;

// The sample creates a workbook, defines a master smart marker that repeats for each department row, and nests employee smart markers that repeat for employees linked via a DataSet relation. WorkbookDesigner processes the markers and saves the output as SmartMarkersNested.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ----- Master smart marker for each department -----
        // This marker will repeat for each row in the Departments table
        sheet.Cells["A1"].PutValue("Department: &=\"DepartmentName\"");

        // ----- Header for nested employee details -----
        sheet.Cells["A3"].PutValue("Employee Name");
        sheet.Cells["B3"].PutValue("Title");

        // ----- Employee detail smart markers (nested) -----
        // These markers will repeat for each employee belonging to the current department
        sheet.Cells["A4"].PutValue("&=\"EmployeeName\"");
        sheet.Cells["B4"].PutValue("&=\"EmployeeTitle\"");

        // ----- Prepare data source with master-detail relationship -----
        DataSet dataSet = new DataSet();

        // Master table: Departments
        DataTable deptTable = new DataTable("Departments");
        deptTable.Columns.Add("DepartmentID", typeof(int));
        deptTable.Columns.Add("DepartmentName", typeof(string));
        deptTable.Rows.Add(1, "Sales");
        deptTable.Rows.Add(2, "Engineering");
        dataSet.Tables.Add(deptTable);

        // Detail table: Employees
        DataTable empTable = new DataTable("Employees");
        empTable.Columns.Add("DepartmentID", typeof(int));
        empTable.Columns.Add("EmployeeName", typeof(string));
        empTable.Columns.Add("EmployeeTitle", typeof(string));
        empTable.Rows.Add(1, "Alice", "Sales Manager");
        empTable.Rows.Add(1, "Bob", "Sales Representative");
        empTable.Rows.Add(2, "Charlie", "Software Engineer");
        empTable.Rows.Add(2, "Diana", "Senior Engineer");
        dataSet.Tables.Add(empTable);

        // Define the relation between Departments and Employees
        dataSet.Relations.Add("DeptEmp",
            deptTable.Columns["DepartmentID"],
            empTable.Columns["DepartmentID"]);

        // ----- Process smart markers -----
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dataSet);
        designer.Process();

        // ----- Save the resulting workbook -----
        workbook.Save("SmartMarkersNested.xlsx");
    }
}
