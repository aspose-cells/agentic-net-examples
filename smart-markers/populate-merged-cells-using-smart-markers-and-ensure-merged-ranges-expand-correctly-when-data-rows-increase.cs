// Title: How to populate merged cells with smart markers and auto‑expand rows using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that merges a header row and a data row, inserts smart markers (&=Header, &=Name, &=Age), binds a DataTable, and calls WorkbookDesigner.Process to repeat the merged range for each record. | Show an example of using WorkbookDesigner with a DataTable to automatically expand merged cells when generating an Excel report with Aspose.Cells. | Generate a complete C# program that creates a workbook, defines merged cells with smart markers, processes the template, and saves the output file.
// Common Searches: asp.net merge cells and use smart markers to repeat rows in Excel | Aspose.Cells expand merged range for each DataTable row | C# smart markers with merged header and data rows example | populate Excel template with merged cells from DataTable using Aspose.Cells | auto‑expand merged rows in Excel using WorkbookDesigner and smart markers
// Tags: smart markers merged cells Aspose.Cells | auto expand merged rows Aspose.Cells | WorkbookDesigner populate merged cells C# | DataTable to Excel smart markers | merged header template Aspose.Cells | C# Aspose.Cells repeat merged range

using System;
using System.Data;
using Aspose.Cells;

// // Demonstrates creating a workbook, merging header and data row cells, placing smart markers, binding an employee DataTable, processing the markers so the merged range expands for each record, and saving the result as an Excel file.
class SmartMarkerMergedCellsExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Define the template with merged cells and smart markers
        // -------------------------------------------------

        // Merge cells A1:B1 for the header (will expand automatically)
        sheet.Cells.Merge(0, 0, 1, 2); // Row 0, Column 0, 1 row, 2 columns (A1:B1)
        // Place a smart marker in the merged header cell
        sheet.Cells["A1"].PutValue("&=Header");

        // Merge cells A2:B2 for the data row template
        sheet.Cells.Merge(1, 0, 1, 2); // Row 1, Column 0, 1 row, 2 columns (A2:B2)
        // Put smart markers inside the merged range
        // The first marker will be used for the Name column, the second for Age
        sheet.Cells["A2"].PutValue("&=Name");
        sheet.Cells["B2"].PutValue("&=Age");

        // -------------------------------------------------
        // Prepare the data source (DataTable) with multiple rows
        // -------------------------------------------------
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Age", typeof(int));

        // Add sample rows
        dt.Rows.Add("Alice", 30);
        dt.Rows.Add("Bob", 45);
        dt.Rows.Add("Charlie", 28);
        dt.Rows.Add("Diana", 35);

        // -------------------------------------------------
        // Process smart markers
        // -------------------------------------------------
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);
        designer.Process(); // This expands the merged range for each data row

        // -------------------------------------------------
        // Save the result
        // -------------------------------------------------
        workbook.Save("SmartMarkerMergedCells.xlsx");
    }
}
