// Title: Generate a master‑detail Excel report with nested smart markers for departments and employees using Aspose.Cells for .NET
// AI Prompts: Design an Excel template that places a master smart marker '&=Departments' to repeat department rows and embeds employee smart markers '&=Employees.EmpName' and '&=Employees.Salary' inside each repeated block. | Create a DataSet with Departments and Employees tables, define a DeptID relation, bind it to WorkbookDesigner, process the smart markers, and save the resulting workbook.
// Common Searches: aspocells c# master smart marker repeat rows for each department | nested employee smart markers inside department block Aspose.Cells example | how to use DataSet relations with WorkbookDesigner for hierarchical Excel output | c# generate master‑detail Excel file using smart markers and dataset | aspocells smart marker master‑detail implementation tutorial
// Tags: master smart marker repeat rows Aspose.Cells | nested detail smart markers WorkbookDesigner | dataset relation hierarchical Excel export | c# master‑detail smart marker example | aspocells generate Excel with parent‑child data

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerMasterDetailDemo
{
    // Shows how to build an Excel template that uses a master smart marker to repeat department rows and nests employee smart markers for each department, leveraging a DataSet with a DeptID relation and WorkbookDesigner to produce a master‑detail spreadsheet.
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and obtain the first sheet
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 2. Build the template with smart markers
            // -------------------------------------------------
            // Header row
            cells["A1"].PutValue("Department");
            cells["B1"].PutValue("Employee");
            cells["C1"].PutValue("Salary");

            // Row that starts the repeat for the master table (Departments)
            // The marker "&=Departments" tells Aspose.Cells to repeat the following rows for each department row.
            cells["A2"].PutValue("&=Departments");

            // Inside the repeated block we place master column marker and start the detail repeat (Employees)
            cells["A3"].PutValue("&=Departments.DeptName");          // Master column
            cells["B3"].PutValue("&=Employees.EmpName");            // Detail column
            cells["C3"].PutValue("&=Employees.Salary");             // Detail column

            // -------------------------------------------------
            // 3. Prepare the data source (DataSet with relation)
            // -------------------------------------------------
            DataSet ds = new DataSet();

            // Master table: Departments
            DataTable dtDept = new DataTable("Departments");
            dtDept.Columns.Add("DeptID", typeof(int));
            dtDept.Columns.Add("DeptName", typeof(string));
            dtDept.Rows.Add(1, "Sales");
            dtDept.Rows.Add(2, "HR");
            dtDept.Rows.Add(3, "IT");
            ds.Tables.Add(dtDept);

            // Detail table: Employees
            DataTable dtEmp = new DataTable("Employees");
            dtEmp.Columns.Add("EmpID", typeof(int));
            dtEmp.Columns.Add("DeptID", typeof(int));
            dtEmp.Columns.Add("EmpName", typeof(string));
            dtEmp.Columns.Add("Salary", typeof(double));
            dtEmp.Rows.Add(101, 1, "John", 50000);
            dtEmp.Rows.Add(102, 1, "Alice", 52000);
            dtEmp.Rows.Add(201, 2, "Bob", 48000);
            dtEmp.Rows.Add(301, 3, "Charlie", 60000);
            dtEmp.Rows.Add(302, 3, "Diana", 62000);
            ds.Tables.Add(dtEmp);

            // Define relation between Departments and Employees on DeptID
            ds.Relations.Add("Dept_Employees",
                dtDept.Columns["DeptID"],
                dtEmp.Columns["DeptID"]);

            // -------------------------------------------------
            // 4. Process the smart markers with WorkbookDesigner
            // -------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;
            designer.SetDataSource(ds);
            designer.Process(); // Populate the template

            // -------------------------------------------------
            // 5. Save the result workbook
            // -------------------------------------------------
            workbook.Save("MasterDetailSmartMarkerOutput.xlsx");
        }
    }
}
