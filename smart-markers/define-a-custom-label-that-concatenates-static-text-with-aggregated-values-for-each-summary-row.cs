using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsCustomTotalsLabel
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Item1");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["A3"].PutValue("Item2");
            worksheet.Cells["B3"].PutValue(150);
            worksheet.Cells["A4"].PutValue("Item3");
            worksheet.Cells["B4"].PutValue(200);

            // Add a table (ListObject) that includes the data range
            int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.ShowTotals = true; // Enable the totals row

            // Configure the totals calculation for the Price column (index 1)
            ListColumn priceColumn = table.ListColumns[1];
            priceColumn.TotalsCalculation = TotalsCalculation.Sum;

            // Calculate the sum of the Price column manually
            double sum = 0;
            for (int row = 1; row <= 3; row++) // data rows are 1‑based (excluding header)
            {
                sum += worksheet.Cells[row, 1].DoubleValue; // column B (index 1)
            }

            // Set a custom label that concatenates static text with the aggregated sum
            priceColumn.TotalsRowLabel = $"Grand Total: {sum}";

            // Save the workbook
            workbook.Save("CustomTotalsLabelDemo.xlsx");
        }
    }
}