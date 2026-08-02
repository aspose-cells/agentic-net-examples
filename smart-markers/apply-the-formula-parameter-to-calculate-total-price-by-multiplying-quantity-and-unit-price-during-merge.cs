using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsMergeWithFormula
{
    class Program
    {
        static void Main()
        {
            // 1. Prepare data source with Quantity and UnitPrice columns
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("Product", typeof(string));
            dt.Columns.Add("Quantity", typeof(int));
            dt.Columns.Add("UnitPrice", typeof(double));

            dt.Rows.Add("Apple", 5, 1.20);
            dt.Rows.Add("Banana", 8, 0.75);
            dt.Rows.Add("Cherry", 12, 2.10);

            // 2. Create a new workbook that will serve as the template
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Header row
            ws.Cells["A1"].PutValue("Product");
            ws.Cells["B1"].PutValue("Quantity");
            ws.Cells["C1"].PutValue("Unit Price");
            ws.Cells["D1"].PutValue("Total Price");

            // Template row with smart markers for data merge
            ws.Cells["A2"].PutValue("&=$Product");      // Smart marker for product name
            ws.Cells["B2"].PutValue("&=$Quantity");    // Smart marker for quantity
            ws.Cells["C2"].PutValue("&=$UnitPrice");   // Smart marker for unit price

            // Formula that calculates total price (Quantity * UnitPrice)
            // This formula will be repeated for each merged row
            ws.Cells["D2"].Formula = "=B2*C2";

            // 3. Configure WorkbookDesigner to repeat formulas for each data row
            WorkbookDesigner designer = new WorkbookDesigner(wb);
            designer.SetDataSource(dt);
            designer.RepeatFormulasWithSubtotal = true; // ensures the formula in D2 is copied to all rows

            // 4. Process the template (merge data and copy formulas)
            designer.Process();

            // 5. Calculate all formulas so that Total Price values are populated
            wb.CalculateFormula();

            // 6. Save the resulting workbook
            wb.Save("MergedWithTotalPrice.xlsx");
        }
    }
}