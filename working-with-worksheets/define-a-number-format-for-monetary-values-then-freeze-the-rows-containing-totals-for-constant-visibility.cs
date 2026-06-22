using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // For ListObject and TotalsCalculation

namespace AsposeCellsNumberFormatAndFreeze
{
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

                // Populate sample data (A1:B5)
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Amount");
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(1200.5);
                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(850.75);
                cells["A4"].PutValue("Cherry");
                cells["B4"].PutValue(430.0);
                cells["A5"].PutValue("Date");
                cells["B5"].PutValue(1020.25);

                // Add a table (ListObject) that includes the data range and show the totals row
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.ShowTotals = true;

                // Set the totals calculation for the "Amount" column to Sum (second column, index 1)
                table.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum;

                // Define a monetary number format style (e.g., $#,##0.00)
                Style moneyStyle = workbook.CreateStyle();
                moneyStyle.Custom = "$#,##0.00";

                // Apply only the number format to the entire "Amount" column (B)
                StyleFlag flag = new StyleFlag { NumberFormat = true };
                // Column B is index 1, rows 0 through 5 (including totals row)
                Aspose.Cells.Range amountRange = cells.CreateRange(0, 1, 6, 1);
                amountRange.ApplyStyle(moneyStyle, flag);

                // Determine the row index of the totals row (zero‑based)
                int totalsRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount;

                // Freeze panes so that rows up to and including the totals row stay visible
                sheet.FreezePanes(totalsRowIndex + 1, 0, totalsRowIndex + 1, 0);

                // Save the workbook
                string outputPath = "NumberFormatAndFreezeTotals.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}