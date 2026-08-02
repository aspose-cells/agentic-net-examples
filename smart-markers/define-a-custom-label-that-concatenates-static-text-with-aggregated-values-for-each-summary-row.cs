// Title: Aspose.Cells .NET – Custom Totals Row Labels with Text + SUBTOTAL
// Description: Demonstrates how to create a workbook, add a ListObject table, enable the totals row, and replace the default totals with formulas that prepend static text (e.g., "Total Quantity: ") to SUBTOTAL results. The example also shows applying bold font and a light‑gray background to the totals row before saving the file.
// Keywords: Aspose.Cells | C# | .NET | custom totals label | SUBTOTAL formula | ListObject table | Excel totals row | concatenate text and formula | cell styling | example code
// Common Searches: Aspose.Cells custom totals row label | add static text to SUBTOTAL in Aspose.Cells | C# create totals row with custom text | format totals row Aspose.Cells .NET | concatenate string and formula Excel using Aspose
// Developer Intent: Generate a totals row where each cell combines a descriptive label with an aggregated value calculated by SUBTOTAL.
// Use Cases: Show "Total Products: X" where X is the count of items in the first column. | Display "Total Quantity: Y" where Y is the sum of the Quantity column. | Present "Total Price: Z" where Z is the sum of the Price column, with bold gray styling for the entire row.
// AI Prompts: Write C# code using Aspose.Cells to add a ListObject table, enable a totals row, and set formulas that concatenate a label with SUBTOTAL results for each column. | Provide an example that applies bold font and a light‑gray background to a custom totals row in Aspose.Cells. | Explain how to calculate the data range indices and build SUBTOTAL formula strings for custom totals labels in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsCustomTotalsLabel
{
    // Demonstrates how to create a workbook, add a ListObject table, enable the totals row, and replace the default totals with formulas that prepend static text (e.g., "Total Quantity: ") to SUBTOTAL results. The example also shows applying bold font and a light‑gray background to the totals row before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Quantity");
                worksheet.Cells["C1"].PutValue("Price");

                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["C2"].PutValue(2.5);

                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["C3"].PutValue(1.8);

                worksheet.Cells["A4"].PutValue("Cherry");
                worksheet.Cells["B4"].PutValue(15);
                worksheet.Cells["C4"].PutValue(3.0);

                // Add a table that includes the data range (including header)
                int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 2, true);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.ShowTotals = true; // Enable totals (summary) row

                // Set totals calculation for each column
                table.ListColumns[0].TotalsCalculation = TotalsCalculation.Count; // Count of products
                table.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum;   // Sum of Quantity
                table.ListColumns[2].TotalsCalculation = TotalsCalculation.Sum;   // Sum of Price

                // Determine rows for data and totals
                int dataStartRow = table.DataRange.FirstRow + 1; // first data row (skip header)
                int dataEndRow = table.DataRange.FirstRow + table.DataRange.RowCount - 1;
                int totalsRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount; // row after data

                // Build custom labels with formulas in the totals row
                // Column A – count with custom label
                worksheet.Cells[totalsRowIndex, 0].Formula = $"\"Total Products: \" & SUBTOTAL(103,{CellsHelper.CellIndexToName(dataStartRow, 0)}:{CellsHelper.CellIndexToName(dataEndRow, 0)})";

                // Column B – sum with custom label
                worksheet.Cells[totalsRowIndex, 1].Formula = $"\"Total Quantity: \" & SUBTOTAL(9,{CellsHelper.CellIndexToName(dataStartRow, 1)}:{CellsHelper.CellIndexToName(dataEndRow, 1)})";

                // Column C – sum with custom label
                worksheet.Cells[totalsRowIndex, 2].Formula = $"\"Total Price: \" & SUBTOTAL(9,{CellsHelper.CellIndexToName(dataStartRow, 2)}:{CellsHelper.CellIndexToName(dataEndRow, 2)})";

                // Apply formatting to the totals row for better readability
                Style totalStyle = workbook.CreateStyle();
                totalStyle.Font.IsBold = true;
                totalStyle.ForegroundColor = System.Drawing.Color.LightGray;
                totalStyle.Pattern = BackgroundType.Solid;

                for (int col = 0; col <= 2; col++)
                {
                    worksheet.Cells[totalsRowIndex, col].SetStyle(totalStyle);
                }

                // Save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomTotalsLabelDemo.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
