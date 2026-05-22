using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTotalsRowDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (header + 5 rows, 3 columns)
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Quantity");
            cells["C1"].PutValue("Price");

            for (int i = 2; i <= 6; i++)
            {
                cells[$"A{i}"].PutValue($"Product {i - 1}");
                cells[$"B{i}"].PutValue(i * 10);          // numeric column
                cells[$"C{i}"].PutValue(i * 2.5);         // numeric column
            }

            // Add a table that covers the data range
            int tableIndex = sheet.ListObjects.Add("A1", "C6", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // -----------------------------------------------------------------
            // Detect tables without a totals row and add one with default sums
            // -----------------------------------------------------------------
            foreach (ListObject lo in sheet.ListObjects)
            {
                // If the table already shows a totals row, skip it
                if (lo.ShowTotals)
                    continue;

                // Enable the totals row
                lo.ShowTotals = true;

                // Set each column's TotalsCalculation to Sum (default for numeric data)
                // For the first column (usually a label), set a friendly label instead
                for (int col = 0; col < lo.ListColumns.Count; col++)
                {
                    ListColumn lc = lo.ListColumns[col];
                    if (col == 0)
                    {
                        // First column: display a label like "Total"
                        lc.TotalsRowLabel = "Grand Total";
                        lc.TotalsCalculation = TotalsCalculation.None;
                    }
                    else
                    {
                        // Other columns: sum the values
                        lc.TotalsCalculation = TotalsCalculation.Sum;
                    }
                }
            }

            // Save the workbook
            workbook.Save("WorkbookWithTotalsRows.xlsx");
        }
    }
}