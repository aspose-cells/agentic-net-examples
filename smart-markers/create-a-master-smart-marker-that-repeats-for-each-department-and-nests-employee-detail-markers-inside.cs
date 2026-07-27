using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook that will act as the template
            Workbook template = new Workbook();
            Worksheet sheet = template.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Define the master smart marker that repeats for each department
            //    The marker "&=Departments" tells Aspose.Cells to repeat the rows
            //    until the next empty row for each record in the Departments table.
            cells["A1"].PutValue("&=Departments");
            cells["A2"].PutValue("Department:");
            cells["B2"].PutValue("&=Departments.DepartmentName"); // Department name

            // 3. Add a header for employee details (will be repeated inside each department)
            cells["A4"].PutValue("Employee Name");
            cells["B4"].PutValue("Employee Age");

            // 4. Define nested smart markers for employee details.
            //    The marker "&=Employees" starts a nested repeat block for the Employees table
            //    that is related to the current department.
            cells["A5"].PutValue("&=Employees");
            cells["A5"].PutValue("&=Employees.EmployeeName");
            cells["B5"].PutValue("&=Employees.Age");

            // 5. Prepare hierarchical data using a DataSet with two DataTables and a relation
            DataSet ds = new DataSet();

            // Departments table
            DataTable dtDept = new DataTable("Departments");
            dtDept.Columns.Add("DepartmentID", typeof(int));
            dtDept.Columns.Add("DepartmentName", typeof(string));
            dtDept.Rows.Add(1, "Sales");
            dtDept.Rows.Add(2, "Engineering");
            ds.Tables.Add(dtDept);

            // Employees table
            DataTable dtEmp = new DataTable("Employees");
            dtEmp.Columns.Add("EmployeeID", typeof(int));
            dtEmp.Columns.Add("DepartmentID", typeof(int));
            dtEmp.Columns.Add("EmployeeName", typeof(string));
            dtEmp.Columns.Add("Age", typeof(int));
            dtEmp.Rows.Add(1, 1, "John Doe", 30);
            dtEmp.Rows.Add(2, 1, "Jane Smith", 28);
            dtEmp.Rows.Add(3, 2, "Alice Johnson", 35);
            dtEmp.Rows.Add(4, 2, "Bob Brown", 40);
            ds.Tables.Add(dtEmp);

            // Define relation between Departments and Employees
            DataRelation rel = new DataRelation("DeptEmp",
                dtDept.Columns["DepartmentID"],
                dtEmp.Columns["DepartmentID"]);
            ds.Relations.Add(rel);

            // 6. Initialize WorkbookDesigner with the template workbook
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = template;

            // 7. Set the hierarchical data source
            designer.SetDataSource(ds);

            // 8. Process the smart markers
            designer.Process();

            // 9. Save the result
            designer.Workbook.Save("MasterSmartMarkerOutput.xlsx");
        }
    }
}