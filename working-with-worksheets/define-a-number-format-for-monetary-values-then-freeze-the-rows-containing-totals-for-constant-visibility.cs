// Title: Aspose.Cells .NET: Apply Currency Format to a Table Column and Freeze the Totals Row
// Description: C# example that creates a workbook, builds a ListObject table with item prices, adds a totals row that sums the Price column, defines a custom currency style ($#,##0.00), applies the style only to the Price column (including the totals row), and freezes all rows through the totals row for constant visibility.
// Keywords: Aspose.Cells | C# | .NET | currency number format | custom number format | freeze panes | freeze rows | totals row | ListObject | table column formatting | Excel monetary formatting
// Common Searches: Aspose.Cells format column as currency C# | How to freeze rows up to totals row in Aspose.Cells | Apply custom number format to ListObject column Aspose.Cells | Add totals row and keep it visible in Excel using Aspose.Cells | C# example for currency style and freeze panes with Aspose.Cells
// Developer Intent: Create a worksheet where the price column shows monetary values in a custom currency format and the totals row stays visible while scrolling.
// Use Cases: Sales report with prices displayed as $1,234.56 and a frozen totals row for quick reference. | Financial statement where the sum row remains on screen while reviewing detailed line items. | Invoice workbook that formats all monetary values as currency and locks the totals row in place.
// AI Prompts: Generate C# code using Aspose.Cells to apply a $#,##0.00 number format to a specific ListObject column and freeze rows through the totals row. | Show how to add a totals row to a table, set its calculation to Sum, format the column as currency, and keep the totals row visible with freeze panes in Aspose.Cells .NET. | Provide an Aspose.Cells example that creates a custom currency style, applies it only to the data and totals rows of a column, then freezes those rows for constant visibility.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsNumberFormatAndFreeze
{
    // C# example that creates a workbook, builds a ListObject table with item prices, adds a totals row that sums the Price column, defines a custom currency style ($#,##0.00), applies the style only to the Price column (including the totals row), and freezes all rows through the totals row for constant visibility.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data: Item names in column A and prices in column B
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Price");
                string[] items = { "Book", "Pen", "Notebook", "Eraser" };
                double[] prices = { 12.5, 1.2, 5.75, 0.8 };

                for (int i = 0; i < items.Length; i++)
                {
                    cells[i + 1, 0].PutValue(items[i]);   // Column A
                    cells[i + 1, 1].PutValue(prices[i]); // Column B
                }

                // Add a table (ListObject) that includes the data range (A1:B{items.Length + 1})
                int lastDataRow = items.Length + 1; // includes header row
                int tableIndex = sheet.ListObjects.Add(0, 0, lastDataRow, 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "SalesTable";

                // Show totals row and set the totals calculation for the Price column (index 1)
                table.ShowTotals = true;
                table.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum;

                // Define a monetary number format style (e.g., $1,234.56)
                Style moneyStyle = workbook.CreateStyle();
                moneyStyle.Custom = "$#,##0.00";

                // Apply only the number format to the Price column (including the totals row)
                StyleFlag flag = new StyleFlag { NumberFormat = true };

                // Determine the range that covers the Price column of the table (data + totals)
                int priceColumnIndex = 1; // column B (zero‑based)
                int dataStartRow = 1; // first data row (zero‑based, after header)
                int totalsRowIndex = items.Length + 1; // zero‑based index of the totals row
                int totalRows = items.Length + 1; // data rows + totals row

                // Create the range for the Price column (data rows + totals row)
                Aspose.Cells.Range priceRange = sheet.Cells.CreateRange(dataStartRow, priceColumnIndex, totalRows, 1);
                priceRange.ApplyStyle(moneyStyle, flag);

                // Freeze the rows up to (and including) the totals row so they stay visible while scrolling
                sheet.FreezePanes(totalsRowIndex + 1, 0, totalsRowIndex + 1, 0);

                // Save the workbook
                workbook.Save("NumberFormatAndFreeze.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
