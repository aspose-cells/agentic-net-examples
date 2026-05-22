using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTotalsRowDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate header row
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Quantity");
            cells["C1"].PutValue("Price");

            // Populate some sample data (numeric columns are Quantity and Price)
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(10);
            cells["C2"].PutValue(0.5);

            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(20);
            cells["C3"].PutValue(0.3);

            cells["A4"].PutValue("Orange");
            cells["B4"].PutValue(15);
            cells["C4"].PutValue(0.4);

            // Create a table that includes the data range (A1:C4)
            int tableIndex = worksheet.ListObjects.Add("A1", "C4", true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Enable the totals row for the table
            table.ShowTotals = true;

            // Configure sum calculation for each numeric column
            // Column 1 (Quantity)
            table.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum;
            table.ListColumns[1].TotalsRowLabel = "Total Quantity";

            // Column 2 (Price)
            table.ListColumns[2].TotalsCalculation = TotalsCalculation.Sum;
            table.ListColumns[2].TotalsRowLabel = "Total Price";

            // Optionally, you can set a custom label for the first (non‑numeric) column
            table.ListColumns[0].TotalsRowLabel = "Grand Total";

            // Save the workbook
            workbook.Save("TableWithTotals.xlsx");
        }
    }
}