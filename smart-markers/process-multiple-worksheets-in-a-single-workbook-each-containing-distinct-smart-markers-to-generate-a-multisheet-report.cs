// Title: Generate a multi‑sheet Excel report by processing smart markers on each worksheet with Aspose.Cells for .NET (C#)
// Common Searches: c# workbookdesigner setdatasource for each worksheet smart markers | asp.net generate multi‑sheet Excel report using a smart‑marker template | process smart markers across multiple sheets with Aspose.Cells .NET | assign different DataTables to separate worksheets for smart‑marker filling
// Tags: Aspose.Cells WorkbookDesigner setdatasource per worksheet | process smart markers in all worksheets | generate multi‑sheet Excel report from template | C# DataTable as smart marker source | Excel template with sheet‑specific smart markers

// Load the template workbook that contains smart markers on each worksheet
Aspose.Cells.WorkbookDesigner designer = new Aspose.Cells.WorkbookDesigner();
designer.Workbook = new Aspose.Cells.Workbook("Template.xlsx");

// -------------------------------------------------
// Prepare data for the first worksheet (e.g., Employees)
// -------------------------------------------------
System.Data.DataTable dtEmployees = new System.Data.DataTable("Employees");
dtEmployees.Columns.Add("Name", typeof(string));
dtEmployees.Columns.Add("Age", typeof(int));
dtEmployees.Rows.Add("John Doe", 30);
dtEmployees.Rows.Add("Jane Smith", 28);
dtEmployees.Rows.Add("Bob Johnson", 45);

// -------------------------------------------------
// Prepare data for the second worksheet (e.g., Products)
// -------------------------------------------------
System.Data.DataTable dtProducts = new System.Data.DataTable("Products");
dtProducts.Columns.Add("ProductName", typeof(string));
dtProducts.Columns.Add("Price", typeof(double));
dtProducts.Rows.Add("Pen", 1.5);
dtProducts.Rows.Add("Notebook", 3.0);
dtProducts.Rows.Add("Eraser", 0.75);

// -------------------------------------------------
// Assign data sources to the designer.
// The smart markers on each sheet must reference these names.
// -------------------------------------------------
designer.SetDataSource("Employees", dtEmployees);
designer.SetDataSource("Products", dtProducts);

// -------------------------------------------------
// Process all smart markers across all worksheets.
// -------------------------------------------------
designer.Process();

// -------------------------------------------------
// Save the generated multi‑sheet report.
// -------------------------------------------------
designer.Workbook.Save("Report.xlsx");
