using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerSummary
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create sample data tables that will be bound to smart markers
            // ------------------------------------------------------------
            DataTable dtQ1 = new DataTable("Sales_Q1");
            dtQ1.Columns.Add("Product", typeof(string));
            dtQ1.Columns.Add("Amount", typeof(double));
            dtQ1.Rows.Add("Apple", 1200);
            dtQ1.Rows.Add("Banana", 850);
            dtQ1.Rows.Add("Cherry", 430);

            DataTable dtQ2 = new DataTable("Sales_Q2");
            dtQ2.Columns.Add("Product", typeof(string));
            dtQ2.Columns.Add("Amount", typeof(double));
            dtQ2.Rows.Add("Apple", 1500);
            dtQ2.Rows.Add("Banana", 970);
            dtQ2.Rows.Add("Cherry", 610);

            // Put the tables into a DataSet – each table will populate a separate worksheet
            DataSet ds = new DataSet();
            ds.Tables.Add(dtQ1);
            ds.Tables.Add(dtQ2);

            // ------------------------------------------------------------
            // 2. Load a template workbook that contains smart markers.
            //    For this example we create the template in memory.
            // ------------------------------------------------------------
            Workbook templateWb = new Workbook();
            // Worksheet for Q1
            Worksheet wsQ1 = templateWb.Worksheets[0];
            wsQ1.Name = "Sales_Q1";
            wsQ1.Cells["A1"].PutValue("&=$Product");   // smart marker for product column
            wsQ1.Cells["B1"].PutValue("&=$Amount");    // smart marker for amount column

            // Worksheet for Q2
            Worksheet wsQ2 = templateWb.Worksheets.Add("Sales_Q2");
            wsQ2.Cells["A1"].PutValue("&=$Product");
            wsQ2.Cells["B1"].PutValue("&=$Amount");

            // ------------------------------------------------------------
            // 3. Bind the data source to the designer and process smart markers
            // ------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = templateWb;
            designer.SetDataSource(ds);
            designer.Process(); // uses the rule WorkbookDesigner.Process()

            // ------------------------------------------------------------
            // 4. Add a summary worksheet that aggregates totals from the data sheets
            // ------------------------------------------------------------
            Worksheet summarySheet = designer.Workbook.Worksheets.Add("Summary");
            // Header row
            summarySheet.Cells["A1"].PutValue("Worksheet");
            summarySheet.Cells["B1"].PutValue("Total Amount");

            // Iterate over all worksheets except the newly added summary sheet
            int summaryRow = 1; // zero‑based index (row 2 in Excel)
            foreach (Worksheet ws in designer.Workbook.Worksheets)
            {
                if (ws.Name == "Summary")
                    continue; // skip the summary sheet itself

                // Determine the last row that contains data in the "Amount" column (column B, index 1)
                int lastDataRow = ws.Cells.MaxDataRow; // zero‑based
                // If there is no data, skip
                if (lastDataRow < 1) continue;

                // Write the worksheet name in the summary sheet
                summarySheet.Cells[summaryRow, 0].PutValue(ws.Name);

                // Build the SUM formula: =SUM('SheetName'!B2:B{lastRow+1})
                // Excel rows are 1‑based, so add 1 to the zero‑based indices.
                string amountColumnLetter = CellsHelper.ColumnIndexToName(1); // "B"
                int startRow = 2; // data starts at row 2 (after header)
                int endRow = lastDataRow + 1; // convert to 1‑based
                string formula = $"=SUM('{ws.Name}'!{amountColumnLetter}{startRow}:{amountColumnLetter}{endRow})";

                // Place the formula in the summary sheet
                summarySheet.Cells[summaryRow, 1].Formula = formula;

                summaryRow++;
            }

            // ------------------------------------------------------------
            // 5. Save the resulting workbook
            // ------------------------------------------------------------
            designer.Workbook.Save("SmartMarkerSummary.xlsx");
        }
    }
}