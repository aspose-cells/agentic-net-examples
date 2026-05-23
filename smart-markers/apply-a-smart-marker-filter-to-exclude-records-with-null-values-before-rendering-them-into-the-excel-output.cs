using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // ---------- Create a workbook template ----------
        Workbook workbook = new Workbook();                     // create
        Worksheet sheet = workbook.Worksheets[0];

        // Header row
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Price");

        // Smart markers for data rows
        sheet.Cells["A2"].PutValue("&=$Products.Name");
        sheet.Cells["B2"].PutValue("&=$Products.Price");

        // Mark the range that contains smart markers (required for processing)
        sheet.Cells.CreateRange("A2:B2").Name = "_CellsSmartMarkers";

        // ---------- Prepare data source with null values ----------
        DataTable dt = new DataTable("Products");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Price", typeof(double));

        dt.Rows.Add("Apple", 1.2);
        dt.Rows.Add(DBNull.Value, 2.5);          // Null Name
        dt.Rows.Add("Banana", DBNull.Value);    // Null Price
        dt.Rows.Add("Cherry", 3.0);

        // ---------- Process smart markers ----------
        WorkbookDesigner designer = new WorkbookDesigner(); // create designer
        designer.Workbook = workbook;
        designer.UpdateEmptyStringAsNull = true;           // treat empty strings as null
        designer.SetDataSource(dt);
        designer.Process();                               // process smart markers

        // ---------- Apply AutoFilter to exclude rows with null (blank) values ----------
        // Determine the used range after processing
        int lastRow = sheet.Cells.MaxDataRow + 1; // +1 because rows are zero‑based
        sheet.AutoFilter.Range = $"A1:B{lastRow}";
        // Exclude rows where the Name column (index 0) is blank/null
        sheet.AutoFilter.MatchNonBlanks(0);
        sheet.AutoFilter.Refresh();                       // apply filter

        // ---------- Save the result ----------
        workbook.Save("FilteredOutput.xlsx");              // save
    }
}