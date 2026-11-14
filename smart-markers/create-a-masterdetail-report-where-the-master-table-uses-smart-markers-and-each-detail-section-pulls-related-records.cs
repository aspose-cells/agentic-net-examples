// Title: Create a master‑detail Excel report with hierarchical smart markers using Aspose.Cells WorkbookDesigner in C#
// Common Searches: C# example for generating master detail Excel using Aspose.Cells smart markers | how to use WorkbookDesigner with DataSet relations for hierarchical smart markers | Aspose.Cells hierarchical smart markers for employee orders report | populate Excel master‑detail tables from DataSet in .NET with smart markers | Aspose.Cells master‑detail report with related detail rows using smart markers | generate Excel report with master‑detail relationship using Aspose.Cells C#
// Tags: WorkbookDesigner data binding with DataSet | C# Aspose.Cells master‑detail Excel generation | smart markers relational DataSet Excel | Excel hierarchical report using Aspose.Cells API | Aspose.Cells employee order report workbook designer

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample creates an Excel workbook, defines a master DataTable for employees and a related detail DataTable for orders, establishes a DataRelation, and uses WorkbookDesigner with hierarchical smart markers to populate master and detail rows. After processing, the workbook is saved as MasterDetailReport.xlsx.
class MasterDetailReport
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Report";

            // ---------- Master Table Smart Markers ----------
            // Header row
            sheet.Cells["A1"].PutValue("Employee ID");
            sheet.Cells["B1"].PutValue("Employee Name");
            // Data row using smart markers
            sheet.Cells["A2"].PutValue("{Master.EmployeeID}");
            sheet.Cells["B2"].PutValue("{Master.Name}");

            // ---------- Detail Section Smart Markers ----------
            // Header for detail rows (starts a few rows below master)
            sheet.Cells["A4"].PutValue("Order ID");
            sheet.Cells["B4"].PutValue("Product");
            sheet.Cells["C4"].PutValue("Quantity");
            // Detail rows using hierarchical smart markers
            sheet.Cells["A5"].PutValue("{Master.Details.OrderID}");
            sheet.Cells["B5"].PutValue("{Master.Details.Product}");
            sheet.Cells["C5"].PutValue("{Master.Details.Quantity}");

            // ---------- Prepare Data Source ----------
            DataSet ds = new DataSet();

            // Master table
            DataTable master = new DataTable("Master");
            master.Columns.Add("EmployeeID", typeof(int));
            master.Columns.Add("Name", typeof(string));
            master.Rows.Add(1, "John Doe");
            master.Rows.Add(2, "Jane Smith");
            ds.Tables.Add(master);

            // Detail table
            DataTable details = new DataTable("Details");
            details.Columns.Add("EmployeeID", typeof(int));
            details.Columns.Add("OrderID", typeof(int));
            details.Columns.Add("Product", typeof(string));
            details.Columns.Add("Quantity", typeof(int));
            details.Rows.Add(1, 1001, "Laptop", 2);
            details.Rows.Add(1, 1002, "Mouse", 5);
            details.Rows.Add(2, 2001, "Keyboard", 3);
            ds.Tables.Add(details);

            // Define relation between master and detail tables
            DataColumn parentCol = master.Columns["EmployeeID"];
            DataColumn childCol = details.Columns["EmployeeID"];
            if (parentCol != null && childCol != null)
            {
                ds.Relations.Add("MasterDetails", parentCol, childCol);
            }

            // ---------- Process Smart Markers ----------
            // Use WorkbookDesigner (the recommended API for smart markers)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(ds);   // attach the data source
            designer.Process();           // generate the report

            // ---------- Save the Workbook ----------
            string outputPath = "MasterDetailReport.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
