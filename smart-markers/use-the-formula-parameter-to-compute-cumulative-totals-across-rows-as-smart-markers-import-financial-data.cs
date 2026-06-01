using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // ---------- 1. Prepare sample financial data ----------
        DataTable financialData = new DataTable("Financial");
        financialData.Columns.Add("Date", typeof(DateTime));
        financialData.Columns.Add("Amount", typeof(double));

        financialData.Rows.Add(new DateTime(2023, 1, 1), 1200.0);
        financialData.Rows.Add(new DateTime(2023, 2, 1), 1500.0);
        financialData.Rows.Add(new DateTime(2023, 3, 1), 1800.0);
        financialData.Rows.Add(new DateTime(2023, 4, 1), 2100.0);
        financialData.Rows.Add(new DateTime(2023, 5, 1), 2400.0);

        // ---------- 2. Create a workbook and design the template ----------
        Workbook workbook = new Workbook();                     // create workbook
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Header row
        cells["A1"].PutValue("Date");
        cells["B1"].PutValue("Amount");
        cells["C1"].PutValue("Cumulative Total");

        // Row 2 contains smart markers for data import
        // &=$Date and &=$Amount will be replaced by the data source values
        cells["A2"].PutValue("&=$Date");
        cells["B2"].PutValue("&=$Amount");
        // Formula that will be repeated for each data row to calculate cumulative sum
        // $B$2 is the absolute start of the Amount column, B2 is relative to the current row
        cells["C2"].Formula = "=SUM($B$2:B2)";

        // ---------- 3. Process smart markers ----------
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.RepeatFormulasWithSubtotal = true;            // repeat the formula for each generated row
        designer.SetDataSource(financialData);
        designer.Process();                                    // import data and repeat the formula

        // ---------- 4. Ensure formulas are calculated ----------
        workbook.CalculateFormula();                           // calculate all formulas

        // ---------- 5. (Optional) Verify cumulative totals ----------
        int lastDataRow = cells.MaxDataRow;                    // last row that contains data
        for (int row = 2; row <= lastDataRow; row++)          // rows are 0‑based; row 2 = Excel row 3
        {
            Cell cumCell = cells[row, 2];                     // column C (index 2)
            Console.WriteLine($"Row {row + 1} cumulative total: {cumCell.Value}");
        }

        // ---------- 6. Save the result ----------
        workbook.Save("FinancialCumulative.xlsx");             // save workbook
    }
}