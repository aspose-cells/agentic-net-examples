using System;
using System.Data;
using Aspose.Cells;

class CumulativeTotalsSmartMarkers
{
    static void Main()
    {
        // 1. Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // 2. Set up the template header
        cells["A1"].PutValue("Date");
        cells["B1"].PutValue("Description");
        cells["C1"].PutValue("Amount");
        cells["D1"].PutValue("Cumulative");

        // 3. Insert smart markers for data rows (starting at row 2)
        //    &=$ColumnName tells Aspose.Cells to replace the cell with data from the DataTable.
        cells["A2"].PutValue("&=$Date");
        cells["B2"].PutValue("&=$Description");
        cells["C2"].PutValue("&=$Amount");

        // 4. Formula to compute cumulative total across rows.
        //    $C$2 is an absolute reference to the first amount cell.
        //    C2 is a relative reference that expands as the formula is copied down.
        cells["D2"].Formula = "=SUM($C$2:C2)";

        // 5. Prepare sample financial data in a DataTable.
        DataTable dt = new DataTable("Financial");
        dt.Columns.Add("Date", typeof(DateTime));
        dt.Columns.Add("Description", typeof(string));
        dt.Columns.Add("Amount", typeof(double));

        dt.Rows.Add(new DateTime(2023, 1, 1), "Opening Balance", 1000.0);
        dt.Rows.Add(new DateTime(2023, 1, 5), "Revenue", 2500.0);
        dt.Rows.Add(new DateTime(2023, 1, 10), "Expense", -800.0);
        dt.Rows.Add(new DateTime(2023, 1, 15), "Revenue", 1200.0);
        dt.Rows.Add(new DateTime(2023, 1, 20), "Expense", -500.0);

        // 6. Use WorkbookDesigner to process smart markers.
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);
        designer.Process();

        // 7. Calculate all formulas so that cumulative totals are evaluated.
        workbook.CalculateFormula();

        // 8. Save the result (lifecycle: save)
        workbook.Save("CumulativeTotals.xlsx");
    }
}