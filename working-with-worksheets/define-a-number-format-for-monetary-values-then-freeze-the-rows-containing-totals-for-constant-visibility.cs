// Title: Apply a custom currency number format and freeze the grand total row in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom currency format for cells | freeze panes to keep totals row visible in Excel with Aspose.Cells | apply monetary style to unit price and line total columns Aspose.Cells | calculate formulas after applying currency formatting Aspose.Cells .NET | how to freeze rows above a specific row using Aspose.Cells C#
// Tags: custom currency format Aspose.Cells C# | freeze panes below totals row Aspose.Cells | apply monetary style to Excel cells | set number format for currency in .NET workbook | keep grand total row visible while scrolling

using System;
using Aspose.Cells;

namespace MonetaryFreezeDemo
{
    // The example creates a workbook, fills it with item, quantity, and unit‑price data, computes line totals with formulas, adds a grand‑total row, defines a custom currency style "$#,##0.00" and applies it to the Unit Price and Line Total columns, freezes all rows above the totals row so the grand total stays in view while scrolling, recalculates formulas, and saves the file as MonetaryFreezeDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // -------------------------------------------------
                // 1. Populate sample data (items, quantity, unit price)
                // -------------------------------------------------
                // Header row
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Quantity");
                cells["C1"].PutValue("Unit Price");
                cells["D1"].PutValue("Line Total");

                // Sample rows
                string[] items = { "Apple", "Orange", "Banana", "Grapes" };
                int[] quantities = { 10, 5, 12, 8 };
                double[] unitPrices = { 1.25, 0.99, 0.75, 2.10 };

                for (int i = 0; i < items.Length; i++)
                {
                    int row = i + 1; // Data starts at row 2 (index 1)
                    cells[$"A{row + 1}"].PutValue(items[i]);
                    cells[$"B{row + 1}"].PutValue(quantities[i]);
                    cells[$"C{row + 1}"].PutValue(unitPrices[i]);

                    // Formula for line total = Quantity * Unit Price
                    cells[$"D{row + 1}"].Formula = $"=B{row + 1}*C{row + 1}";
                }

                // -------------------------------------------------
                // 2. Add a totals row that sums the line totals
                // -------------------------------------------------
                int totalsRowIndex = items.Length + 1; // zero‑based index
                int totalsRowNumber = totalsRowIndex + 1; // Excel row number (starts at 1)

                cells[$"C{totalsRowNumber}"].PutValue("Grand Total:");
                cells[$"D{totalsRowNumber}"].Formula = $"=SUM(D2:D{totalsRowNumber - 1})";

                // -------------------------------------------------
                // 3. Define a monetary number format and apply it
                // -------------------------------------------------
                // Create a style with a custom currency format
                Style monetaryStyle = workbook.CreateStyle();
                monetaryStyle.Custom = "$#,##0.00";

                // Apply the style to Unit Price column (C) and Line Total column (D)
                // Apply to all data rows plus the totals row
                for (int row = 1; row <= totalsRowIndex; row++) // rows 2..(totalsRow)
                {
                    // Column C (Unit Price)
                    cells[row, 2].SetStyle(monetaryStyle);
                    // Column D (Line Total)
                    cells[row, 3].SetStyle(monetaryStyle);
                }

                // -------------------------------------------------
                // 4. Freeze the rows containing the totals for constant visibility
                // -------------------------------------------------
                // Freeze all rows above the totals row (i.e., rows 1..totalsRowIndex)
                // FreezePanes(row, column, freezedRows, freezedColumns)
                // We set the split just below the totals row, so the totals row stays visible while scrolling.
                sheet.FreezePanes(totalsRowNumber, 0, totalsRowNumber, 0);

                // -------------------------------------------------
                // 5. Recalculate formulas and save the workbook
                // -------------------------------------------------
                workbook.CalculateFormula();

                string outputPath = "MonetaryFreezeDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Do not hide failures – report them clearly
                Console.Error.WriteLine("An error occurred while creating the workbook:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
