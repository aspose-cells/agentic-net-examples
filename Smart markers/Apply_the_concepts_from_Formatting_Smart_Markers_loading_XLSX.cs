using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerFormattingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook to act as the template.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Employees";

            // Header row.
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["C1"].PutValue("Salary");
            sheet.Cells["D1"].PutValue("Hire Date");

            // Smart markers row.
            sheet.Cells["A2"].PutValue("&=Employees.Name");
            sheet.Cells["B2"].PutValue("&=Employees.Age");
            sheet.Cells["C2"].PutValue("&=Employees.Salary");
            sheet.Cells["D2"].PutValue("&=Employees.HireDate");

            // Apply formatting to the smart marker cells.
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Number = 164; // "$#,##0.00"
            sheet.Cells["C2"].SetStyle(currencyStyle);

            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // "m/d/yyyy"
            sheet.Cells["D2"].SetStyle(dateStyle);

            // Prepare the data source.
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Salary", typeof(double));
            dt.Columns.Add("HireDate", typeof(DateTime));

            dt.Rows.Add("John Doe", 30, 75000.5, new DateTime(2015, 3, 1));
            dt.Rows.Add("Jane Smith", 28, 82000, new DateTime(2017, 7, 15));
            dt.Rows.Add("Bob Johnson", 45, 95000, new DateTime(2010, 11, 20));

            // Process smart markers.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process();

            // Save the result.
            workbook.Save("SmartMarkersFormattedOutput.xlsx");
        }
    }
}