using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerRunningTotal
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Prepare sample data source (DataTable) with Item and Amount
            // ------------------------------------------------------------
            DataTable dt = new DataTable("Sales");
            dt.Columns.Add("Item", typeof(string));
            dt.Columns.Add("Amount", typeof(double));

            dt.Rows.Add("A", 100);
            dt.Rows.Add("B", 150);
            dt.Rows.Add("C", 200);
            dt.Rows.Add("D", 120);
            dt.Rows.Add("E", 80);

            // ------------------------------------------------------------
            // 2. Create a new workbook and set up the template with smart markers
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Amount");
            cells["C1"].PutValue("Running Total");

            // Row 2 – template row that will be repeated for each data row
            // Smart markers for Item and Amount
            cells["A2"].PutValue("&=$Item");
            cells["B2"].PutValue("&=$Amount");

            // Formula for running total:
            //   =B2 + IF(ROW()>2, C1, 0)
            // For the first data row (row 2) the IF part returns 0,
            // for subsequent rows it adds the previous row's running total (C1, C2, ...).
            cells["C2"].PutValue("=B2+IF(ROW()>2,C1,0)");

            // ------------------------------------------------------------
            // 3. Configure WorkbookDesigner to process smart markers
            // ------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            // Ensure the formula in the template row is repeated for each data row
            designer.RepeatFormulasWithSubtotal = true;
            designer.SetDataSource(dt);

            // Process the smart markers – this will populate rows and copy the formula
            designer.Process();

            // ------------------------------------------------------------
            // 4. Save the resulting workbook
            // ------------------------------------------------------------
            workbook.Save("RunningTotalSmartMarker.xlsx");
        }
    }
}