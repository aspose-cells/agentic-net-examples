// Title: Insert a SUBTOTAL formula and freeze header rows in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# program that adds a SUBTOTAL(9,…) formula to column D, generates grouped subtotals by region, and freezes the top rows using Aspose.Cells. | Use the Cells.Subtotal method to sum sales per region, then call Worksheet.FreezePanes so the header and first subtotal stay visible in the generated Excel file. | Generate an Excel report with automatic subtotals and frozen header rows in C# without relying on Excel interop.
// Common Searches: aspnet add subtotal formula and freeze panes in generated Excel using Aspose.Cells | c# Aspose.Cells how to group rows by column and keep header fixed | programmatically create Excel subtotal rows and freeze top rows with Aspose.Cells .NET | using Aspose.Cells Subtotal method and FreezePanes together in C# example | generate Excel report with region totals and frozen header using Aspose.Cells
// Tags: Aspose.Cells SUBTOTAL function C# | Aspose.Cells FreezePanes API | automatic subtotals grouped by column Aspose.Cells | freeze header rows after subtotal Aspose.Cells | C# generate Excel with subtotal and frozen panes

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalAndFreezeDemo
{
    // The example creates a new workbook, fills it with region, product, and sales data, inserts a SUBTOTAL(9,…) formula in column D, applies automatic subtotals grouped by the Region column, freezes the header and first subtotal row using FreezePanes, and saves the result as SubtotalAndFreezeDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ----- Populate sample data -----
            // Header row
            cells["A1"].PutValue("Region");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            // Data rows
            object[,] data = new object[,]
            {
                {"North", "Widget", 5000},
                {"North", "Gadget", 3000},
                {"South", "Widget", 6000},
                {"South", "Gadget", 4000},
                {"West",  "Widget", 4500}
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                cells[i + 1, 0].PutValue(data[i, 0]); // Region
                cells[i + 1, 1].PutValue(data[i, 1]); // Product
                cells[i + 1, 2].PutValue(data[i, 2]); // Sales
            }

            // ----- Insert a formula that calculates a subtotal for the Sales column -----
            // The SUBTOTAL function (function_num 9 = SUM) will sum the visible cells in the range.
            // Place the formula in column D, row 2 (first data row). It will be copied down later if needed.
            cells["D2"].Formula = "=SUBTOTAL(9,C2:C6)";

            // ----- Add automatic subtotals using the Cells.Subtotal method -----
            // Define the range that includes the header and data (A1:C6)
            CellArea area = CellArea.CreateCellArea("A1", "C6");

            // Group by the first column (Region), sum the Sales column (index 2),
            // replace existing subtotals, no page breaks, place summary rows below data.
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, true, false, true);

            // ----- Freeze panes so that the header and the first subtotal row stay visible -----
            // After Subtotal is applied, the first subtotal row appears after the first group.
            // Freezing at cell "A2" keeps the header row (row 1) fixed while scrolling.
            sheet.FreezePanes("A2", 1, 0);

            // ----- Save the workbook -----
            workbook.Save("SubtotalAndFreezeDemo.xlsx");
        }
    }
}
