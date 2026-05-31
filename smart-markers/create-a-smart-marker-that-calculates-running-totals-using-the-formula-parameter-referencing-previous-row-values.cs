using System;
using System.Data;
using Aspose.Cells;

class RunningTotalSmartMarker
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // ----- Template setup -----
        // Header row
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Amount");
        cells["C1"].PutValue("Running Total");

        // Row 2 contains the smart markers and the running‑total formula
        // &=$Item and &=$Amount are smart markers that will be replaced by data source values
        cells["A2"].PutValue("&=$Item");
        cells["B2"].PutValue("&=$Amount");
        // Running total: sum of current amount (B2) and the total from the previous row (C1)
        // When the formula is repeated, relative references will adjust automatically:
        // Row 3 -> =SUM(B3, C2), Row 4 -> =SUM(B4, C3), etc.
        cells["C2"].Formula = "=SUM(B2, C1)";

        // ----- Data source -----
        DataTable dt = new DataTable("Data");
        dt.Columns.Add("Item", typeof(string));
        dt.Columns.Add("Amount", typeof(double));
        dt.Rows.Add("Item A", 120);
        dt.Rows.Add("Item B", 80);
        dt.Rows.Add("Item C", 150);
        dt.Rows.Add("Item D", 200);

        // ----- Designer processing -----
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        // Ensure that the formula in C2 is repeated for each data row
        designer.RepeatFormulasWithSubtotal = true;
        designer.SetDataSource(dt);
        designer.Process();

        // ----- Save the result -----
        workbook.Save("RunningTotalSmartMarker.xlsx");
    }
}