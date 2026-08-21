// Title: Aspose.Cells for .NET – Create a custom totals‑row label that includes the summed value (C#)
// Description: Demonstrates how to build a workbook, add a ListObject table, enable the totals row, calculate the sum of the "Price" column, read the computed total, and set a custom TotalsRowLabel that concatenates static text (e.g., "Grand Total") with the aggregated amount, then save the file as an .xlsx document.
// Keywords: Aspose.Cells custom totals row label | C# TotalsRowLabel | ListObject totals calculation | concatenate static text with sum | retrieve table total Aspose.Cells | Excel table Grand Total label | .NET Excel aggregation | Aspose.Cells TotalsCalculation.Sum
// Common Searches: Aspose.Cells set custom totals row label C# | How to read sum from ListObject totals row | Combine text with calculated total in Excel using Aspose.Cells | C# TotalsRowLabel with dynamic value | Aspose.Cells table totals row custom text
// Developer Intent: The developer needs to display a dynamic label in the totals row that merges a fixed phrase with the column’s calculated sum.
// Use Cases: Financial dashboards that show "Grand Total (1234)" in the totals row. | Automated invoice generation where the totals row contains a custom message with the total amount. | Report templates that embed aggregated values inside descriptive labels for clearer presentation.
// AI Prompts: Generate C# code using Aspose.Cells to set TotalsRowLabel to "Grand Total (value)" where value is the sum of a column. | Explain how to fetch the computed total from a ListObject totals row and embed it in a formatted label string. | Show how to update the custom totals row label after modifying data in the worksheet with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsCustomTotalsLabel
{
    // Demonstrates how to build a workbook, add a ListObject table, enable the totals row, calculate the sum of the "Price" column, read the computed total, and set a custom TotalsRowLabel that concatenates static text (e.g., "Grand Total") with the aggregated amount, then save the file as an .xlsx document.
    public class Program
    {
        public static void Main()
        {
            try
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

                // Add a table that includes the header and data rows
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
                int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.ShowTotals = true; // Enable the totals row

                // Configure the totals calculation for the "Price" column (index 1)
                ListColumn priceColumn = table.ListColumns[1];
                priceColumn.TotalsCalculation = TotalsCalculation.Sum;

                // Determine the totals row index (zero‑based)
                int totalsRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount;
                // Determine the column index for "Price"
                int priceColumnIndex = table.DataRange.FirstColumn + 1;

                // Retrieve the computed total value from the totals row
                double totalValue = worksheet.Cells[totalsRowIndex, priceColumnIndex].DoubleValue;

                // Set a custom label that includes the aggregated total
                priceColumn.TotalsRowLabel = $"Grand Total ({totalValue})";

                // Save the workbook
                workbook.Save("CustomTotalsLabelDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
