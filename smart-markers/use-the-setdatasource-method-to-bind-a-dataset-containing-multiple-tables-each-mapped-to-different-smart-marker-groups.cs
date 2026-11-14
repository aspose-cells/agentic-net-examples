// Title: Using WorkbookDesigner.SetDataSource to bind multiple DataSet tables to separate smart marker groups in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# bind DataSet with Employees and Departments tables to smart markers | WorkbookDesigner SetDataSource multiple tables example | How to map different smart marker groups to separate DataTables in Aspose.Cells | Generate Excel report from two DataTables using smart markers in .NET
// Tags: WorkbookDesigner SetDataSource multiple tables | smart markers multiple DataSet tables | C# Aspose.Cells bind DataSet to smart markers | Excel report generation from Employees and Departments | smart marker groups per DataTable

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel template containing smart markers for Employees and Departments, creates a DataSet with two tables, assigns it to a WorkbookDesigner via SetDataSource, processes the markers, and saves the populated report as an Excel file.
class SmartMarkerExample
{
    static void Main()
    {
        try
        {
            // Path to the Excel template that contains smart markers.
            const string templatePath = "SmartMarkerTemplate.xlsx";
            Workbook workbook;

            // Load the template if it exists; otherwise create a simple workbook with sample smart markers.
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                // Sample smart markers for Employees table.
                ws.Cells["A1"].PutValue("${Employees.Name}");
                ws.Cells["B1"].PutValue("${Employees.Age}");
                ws.Cells["C1"].PutValue("${Employees.Salary}");
                // Sample smart markers for Departments table.
                ws.Cells["A5"].PutValue("${Departments.DeptName}");
                ws.Cells["B5"].PutValue("${Departments.Location}");
            }

            // Build a DataSet with Employees and Departments tables.
            DataSet ds = new DataSet();

            DataTable employees = new DataTable("Employees");
            employees.Columns.Add("Name", typeof(string));
            employees.Columns.Add("Age", typeof(int));
            employees.Columns.Add("Salary", typeof(double));
            employees.Rows.Add("John Doe", 30, 55000);
            employees.Rows.Add("Jane Smith", 28, 62000);
            ds.Tables.Add(employees);

            DataTable departments = new DataTable("Departments");
            departments.Columns.Add("DeptName", typeof(string));
            departments.Columns.Add("Location", typeof(string));
            departments.Rows.Add("HR", "New York");
            departments.Rows.Add("IT", "San Francisco");
            ds.Tables.Add(departments);

            // Use WorkbookDesigner to process smart markers.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(ds);
            designer.Process();

            // Save the generated report.
            const string outputPath = "SmartMarkerReport.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
