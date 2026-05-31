using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data (header + rows)
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Quantity");
        cells["C1"].PutValue("Price");

        cells["A2"].PutValue("Apple");
        cells["B2"].PutValue(10);
        cells["C2"].PutValue(2.5);

        cells["A3"].PutValue("Banana");
        cells["B3"].PutValue(5);
        cells["C3"].PutValue(3.0);

        // Add a table that covers the data range
        int tableIndex = worksheet.ListObjects.Add("A1", "C3", true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Initially show the totals row (optional, just to have one)
        table.ShowTotals = true;

        // ----- Remove the existing totals row -----
        table.ShowTotals = false;   // hides/removes the totals row

        // ----- Re‑add the totals row with custom formulas -----
        table.ShowTotals = true;    // shows a new totals row

        // Iterate through each column of the table
        foreach (ListColumn column in table.ListColumns)
        {
            // The first data cell of the column (row index 1 because row 0 is the header)
            Cell firstDataCell = column.Range[1, 0];

            // If the first data cell contains a numeric value, treat the column as numeric
            if (firstDataCell != null && firstDataCell.Type == CellValueType.IsNumeric)
            {
                // Use a custom totals calculation
                column.TotalsCalculation = TotalsCalculation.Custom;

                // Example custom formula: sum of the column values
                string customFormula = $"=SUM([{column.Name}])";

                // Set the custom formula for the totals row of this column
                // Parameters: formula string, isR1C1 = false, isLocal = false
                column.SetCustomTotalsRowFormula(customFormula, false, false);
            }
            else
            {
                // For non‑numeric columns, you can set a label or leave it empty
                column.TotalsCalculation = TotalsCalculation.None;
                column.TotalsRowLabel = "Total";
            }
        }

        // Save the workbook with the updated table totals
        workbook.Save("TableWithCustomTotals.xlsx");
    }
}