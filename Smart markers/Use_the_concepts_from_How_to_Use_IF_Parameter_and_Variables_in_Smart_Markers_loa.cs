using System;
using System.Data;
using Aspose.Cells;

class SmartMarkerIfVariableDemo
{
    static void Main()
    {
        // Create a template workbook with smart markers
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Header row
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["C1"].PutValue("Department");
        sheet.Cells["D1"].PutValue("IfHR");

        // Smart marker rows (the data will be populated from the DataTable)
        sheet.Cells["A2"].PutValue("&=Employees.Name");
        sheet.Cells["B2"].PutValue("&=Employees.Age");
        sheet.Cells["C2"].PutValue("&=Employees.Department");
        // IF smart marker using a variable: if ShowHROnly is true, display the department, otherwise leave blank
        sheet.Cells["D2"].PutValue("&IF(ShowHROnly, =Employees.Department, \"\")");

        // Initialize WorkbookDesigner and assign the workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // Prepare a DataTable as data source
        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("Name", typeof(string));
        employees.Columns.Add("Age", typeof(int));
        employees.Columns.Add("Department", typeof(string));

        employees.Rows.Add("John Doe", 30, "Sales");
        employees.Rows.Add("Jane Smith", 45, "HR");
        employees.Rows.Add("Bob Johnson", 28, "IT");

        // Bind the DataTable to the designer
        designer.SetDataSource(employees);

        // Set a variable that will be used inside the IF smart marker
        designer.SetDataSource("ShowHROnly", true); // Change to false to hide department values

        // Process all smart markers
        designer.Process();

        // Save the processed workbook
        workbook.Save("ResultWithIfAndVariables.xlsx");
    }
}