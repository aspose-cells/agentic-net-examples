// Title: Format a Currency Column and Freeze the Totals Row with Aspose.Cells for .NET
// Description: Creates a workbook, adds a ListObject table with a summed totals row, applies a custom "$#,##0.00" currency style only to the data and totals cells, freezes all rows up to the totals row, and saves the file.
// Keywords: Aspose.Cells currency format | custom number format .NET | freeze panes Aspose.Cells | totals row formatting | ListObject monetary style | C# Excel currency formatting | freeze rows with totals
// Common Searches: Aspose.Cells set custom monetary format for a column | How to freeze rows that include a totals row in Aspose.Cells | Apply number format only to data rows in a ListObject | Add and format a totals row in an Aspose.Cells table | C# freeze panes up to a specific row in Excel
// Developer Intent: Generate a worksheet, format the Price column as currency, add a summed totals row, and keep those rows visible by freezing them.
// Use Cases: Sales report where prices are shown in a specific currency format and the total stays on screen while scrolling. | Invoice template that formats all monetary values with "$#,##0.00" and pins the header and totals rows for quick reference. | Financial dashboard that uses a ListObject with a sum totals row, applies currency styling only to numeric cells, and freezes the totals row.
// AI Prompts: Provide C# code using Aspose.Cells to apply a "$#,##0.00" format to a table column and freeze rows up to the totals row. | Show an Aspose.Cells example that creates a ListObject, adds a sum totals row, formats the numeric column as currency, and freezes the header and totals rows. | Explain how to use StyleFlag in Aspose.Cells to apply only number formatting while preserving other cell styles.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a ListObject table with a summed totals row, applies a custom "$#,##0.00" currency style only to the data and totals cells, freezes all rows up to the totals row, and saves the file.
    public class MonetaryFormatAndFreezeRows
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (Item and Price)
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Price");
            string[] items = { "Apple", "Banana", "Cherry", "Date", "Elderberry" };
            double[] prices = { 1.25, 0.75, 2.50, 3.10, 4.20 };
            for (int i = 0; i < items.Length; i++)
            {
                cells[i + 1, 0].PutValue(items[i]);   // Column A
                cells[i + 1, 1].PutValue(prices[i]); // Column B
            }

            // Add a table covering the data range and enable the totals row
            int tableIndex = worksheet.ListObjects.Add(0, 0, items.Length, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.ShowTotals = true;
            // Set the totals calculation for the Price column to Sum
            table.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum;

            // Define a custom number format for monetary values
            Style moneyStyle = workbook.CreateStyle();
            moneyStyle.Custom = "$#,##0.00";

            // Apply only the number format to the data column (excluding header and totals)
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Data rows are from row 2 to the row before the totals row
            int dataStartRow = 1; // zero‑based index (row 2 in Excel)
            int dataRowCount = items.Length; // number of data rows
            Aspose.Cells.Range dataRange = cells.CreateRange(dataStartRow, 1, dataRowCount, 1);
            dataRange.ApplyStyle(moneyStyle, flag);

            // Calculate the index of the totals row (header + data rows)
            int totalRow = items.Length + 1; // zero‑based index

            // Apply the same number format to the totals row cell
            cells[totalRow, 1].SetStyle(moneyStyle);

            // Freeze rows up to and including the totals row so they stay visible while scrolling
            int rowsToFreeze = totalRow + 1; // number of rows to keep frozen
            worksheet.FreezePanes(rowsToFreeze, 0, rowsToFreeze, 0);

            // Save the workbook
            workbook.Save("MonetaryFormatAndFreezeRows.xlsx");
        }
    }
}
