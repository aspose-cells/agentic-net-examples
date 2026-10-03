// Title: Assign a DataTable to WorkbookDesigner as a custom data source and generate an Excel report with smart markers using Aspose.Cells for .NET
// AI Prompts: Write C# that creates a DataTable, sets it as the data source of a WorkbookDesigner via SetDataSource, processes smart markers, and saves the workbook as an XLSX file. | Demonstrate loading an existing Excel template, inserting a smart‑marker, binding a DataTable to WorkbookDesigner, and outputting the populated report with Aspose.Cells. | Show a try‑catch pattern for assigning a DataTable to WorkbookDesigner and processing smart markers in a .NET console application.
// Common Searches: asp.net how to bind a datatable to workbookdesigner for smart markers | c# aspose.cells populate excel template from custom datatable | set data source for workbookdesigner using setdatasource method | process smart markers with datatable in asp.net core | generate excel report from datatable using aspose.cells workbookdesigner
// Tags: WorkbookDesigner SetDataSource DataTable | smart markers populate Excel from DataTable | Aspose.Cells load template process smart markers | C# generate Excel report with WorkbookDesigner | Excel template binding custom tabular data Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The program builds a DataTable of employee records, loads or creates an Excel workbook, places a smart‑marker, assigns the DataTable to WorkbookDesigner via SetDataSource, processes the marker to fill the sheet, and saves the result as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create a DataTable with custom data ----------
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Salary", typeof(double));

            dt.Rows.Add(1, "John Doe", 50000);
            dt.Rows.Add(2, "Jane Smith", 60000);
            dt.Rows.Add(3, "Bob Johnson", 55000);

            // ---------- Load a template if it exists; otherwise create a new workbook ----------
            string templatePath = "Template.xlsx";
            Workbook workbook = File.Exists(templatePath) ? new Workbook(templatePath) : new Workbook();

            // ---------- Add a worksheet and place a Designer marker ----------
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Report";
            sheet.Cells["A2"].PutValue("&=Employees"); // marker for binding

            // ---------- Assign the DataTable to WorkbookDesigner ----------
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt); // bind the DataTable

            // ---------- Process the designer to populate the worksheet ----------
            designer.Process();

            // ---------- Save the resulting workbook ----------
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
