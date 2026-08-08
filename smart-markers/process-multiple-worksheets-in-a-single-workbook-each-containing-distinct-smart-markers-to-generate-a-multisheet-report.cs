// Title: Create a multi‑sheet Excel report with smart markers using Aspose.Cells for .NET (C#)
// Description: This example builds a workbook containing two worksheets—Employees and Products—adds range smart markers to each sheet, binds separate DataTable sources, runs WorkbookDesigner to populate the markers, and saves the result as a single Excel file.
// Keywords: Aspose.Cells C# smart markers | multi‑sheet Excel report .NET | WorkbookDesigner multiple data sources | range smart markers across worksheets | Aspose.Cells tutorial USA | Aspose.Cells Europe example | C# generate Excel with smart markers
// Common Searches: asp.net generate Excel with smart markers on several sheets | how to bind multiple DataTables to WorkbookDesigner | create employee and product tabs in one Excel file using Aspose.Cells | process smart markers in all worksheets C# | Aspose.Cells multi‑worksheet report tutorial
// Developer Intent: Populate several worksheets in one workbook, each driven by its own smart‑marker layout and DataTable, to produce a consolidated Excel report.
// Use Cases: Combine an employee directory and a product catalog into a single downloadable workbook. | Produce separate monthly KPI tabs (sales, inventory, finance) from distinct data feeds. | Automate generation of multi‑section financial statements where each section uses a dedicated data source.
// AI Prompts: Add a third worksheet for an Orders table with smart markers and show the updated C# code. | Show how to apply currency formatting to price cells on the Products sheet using smart markers. | Generate error‑handling logic for missing DataTable names when processing smart markers across multiple sheets.

using System;
using System.Data;
using Aspose.Cells;

// This example builds a workbook containing two worksheets—Employees and Products—adds range smart markers to each sheet, binds separate DataTable sources, runs WorkbookDesigner to populate the markers, and saves the result as a single Excel file.
class MultiSheetSmartMarkerReport
{
    static void Main()
    {
        // Create a new workbook that will serve as the template
        Workbook wb = new Workbook();

        // ---------- Worksheet 1 : Employees ----------
        Worksheet wsEmployees = wb.Worksheets[0];
        wsEmployees.Name = "Employees";

        // Header row
        wsEmployees.Cells["A1"].PutValue("Name");
        wsEmployees.Cells["B1"].PutValue("Age");

        // Smart markers (range smart markers are used)
        wsEmployees.Cells["A2"].PutValue("&=Employees.Name");
        wsEmployees.Cells["B2"].PutValue("&=Employees.Age");

        // ---------- Worksheet 2 : Products ----------
        Worksheet wsProducts = wb.Worksheets.Add("Products");

        // Header row
        wsProducts.Cells["A1"].PutValue("Product");
        wsProducts.Cells["B1"].PutValue("Price");

        // Smart markers for the second sheet
        wsProducts.Cells["A2"].PutValue("&=Products.ProductName");
        wsProducts.Cells["B2"].PutValue("&=Products.Price");

        // ---------- Prepare data sources ----------
        // DataTable for Employees
        DataTable dtEmployees = new DataTable("Employees");
        dtEmployees.Columns.Add("Name", typeof(string));
        dtEmployees.Columns.Add("Age", typeof(int));
        dtEmployees.Rows.Add("John Doe", 30);
        dtEmployees.Rows.Add("Jane Smith", 28);

        // DataTable for Products
        DataTable dtProducts = new DataTable("Products");
        dtProducts.Columns.Add("ProductName", typeof(string));
        dtProducts.Columns.Add("Price", typeof(double));
        dtProducts.Rows.Add("Laptop", 1200.5);
        dtProducts.Rows.Add("Mouse", 25.99);

        // ---------- Initialize WorkbookDesigner ----------
        WorkbookDesigner designer = new WorkbookDesigner(wb);

        // Set the data sources; the table name is taken from DataTable.TableName
        designer.SetDataSource(dtEmployees);
        designer.SetDataSource(dtProducts);

        // Process all smart markers across all worksheets
        designer.Process();

        // Save the generated multi‑sheet report
        wb.Save("MultiSheetReport.xlsx");
    }
}
