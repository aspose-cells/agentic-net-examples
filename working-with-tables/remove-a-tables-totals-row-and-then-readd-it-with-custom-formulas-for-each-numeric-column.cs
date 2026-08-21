// Title: Aspose.Cells for .NET – Remove and Re‑add Table Totals Row with Custom SUM & AVERAGE Formulas (C#)
// Description: C# sample that creates a workbook, adds a ListObject (Excel table), hides its default totals row, then shows the row again and sets custom formulas – SUM for the Quantity column and AVERAGE for the Price column – before saving the file.
// Keywords: Aspose.Cells | C# | .NET | ListObject | Excel table totals row | remove totals row | custom totals formula | SUM formula Aspose.Cells | AVERAGE formula Aspose.Cells | sample code | GitHub example
// Common Searches: Aspose.Cells hide table totals row C# | set custom totals formula ListObject Aspose.Cells | remove and add totals row Aspose.Cells .NET | custom SUM and AVERAGE in Excel table using Aspose.Cells | C# example for ListObject custom aggregates
// Developer Intent: Hide an existing ListObject totals row, then display it again with user‑defined formulas for numeric columns.
// Use Cases: Generate a sales workbook where the totals row must be rebuilt with specific aggregates after data changes. | Create an Excel export that shows a custom SUM for quantity and an AVERAGE for price in the table footer. | Implement dynamic reporting where the totals row is toggled and customized programmatically.
// AI Prompts: Write C# code with Aspose.Cells to hide a table totals row and then add a custom SUM formula to the Quantity column. | Show how to apply different custom formulas (SUM, AVERAGE, COUNT) to multiple columns of an Aspose.Cells ListObject. | Explain the steps to toggle ShowTotals and set custom totals row formulas without altering other table settings.

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;

// C# sample that creates a workbook, adds a ListObject (Excel table), hides its default totals row, then shows the row again and sets custom formulas – SUM for the Quantity column and AVERAGE for the Price column – before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data (header + 3 rows)
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Quantity");
        cells["C1"].PutValue("Price");
        cells["D1"].PutValue("Notes");

        string[] items = { "Apple", "Banana", "Orange" };
        int[] quantities = { 10, 20, 15 };
        double[] prices = { 0.5, 0.3, 0.6 };
        string[] notes = { "Fresh", "Ripe", "Citrus" };

        for (int i = 0; i < items.Length; i++)
        {
            int row = i + 1; // data starts at row 2 (zero‑based index)
            cells[row, 0].PutValue(items[i]);
            cells[row, 1].PutValue(quantities[i]);
            cells[row, 2].PutValue(prices[i]);
            cells[row, 3].PutValue(notes[i]);
        }

        // Add a table that covers the data range (A1:D4)
        int tableIdx = worksheet.ListObjects.Add(0, 0, items.Length, 3, true);
        ListObject table = worksheet.ListObjects[tableIdx];
        table.DisplayName = "SalesTable";

        // Show the totals row initially
        table.ShowTotals = true;

        // ---- Remove the totals row ----
        table.ShowTotals = false;

        // ---- Re‑add the totals row with custom formulas ----
        table.ShowTotals = true;

        // Column "Quantity" (numeric) – custom SUM formula
        ListColumn qtyColumn = table.ListColumns[1];
        qtyColumn.TotalsCalculation = TotalsCalculation.Custom;
        qtyColumn.SetCustomTotalsRowFormula("=SUM([Quantity])", false, false);

        // Column "Price" (numeric) – custom AVERAGE formula
        ListColumn priceColumn = table.ListColumns[2];
        priceColumn.TotalsCalculation = TotalsCalculation.Custom;
        priceColumn.SetCustomTotalsRowFormula("=AVERAGE([Price])", false, false);

        // Save the workbook
        workbook.Save("TableWithCustomTotals.xlsx");
    }
}
