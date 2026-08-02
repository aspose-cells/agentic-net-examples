// Title: Aspose.Cells for .NET: Remove and Re‑Add a Table Totals Row with Custom Formulas (C#)
// Description: This example demonstrates how to create a workbook, add a ListObject covering A1:D5, hide the existing totals row, show it again, and assign custom formulas to the Qty, Price, and Discount columns using TotalsCalculation.Custom and SetCustomTotalsRowFormula. The workbook is then saved with the customized totals row.
// Keywords: Aspose.Cells | C# | .NET | ListObject | Excel table | Remove totals row | Add totals row | Custom totals formulas | SetCustomTotalsRowFormula | ShowTotals property | TotalsCalculation.Custom | GitHub sample | code example
// Common Searches: how to delete totals row in Aspose.Cells | Aspose.Cells custom totals row formula C# | remove and re‑add table totals row ListObject | set custom totals for numeric columns Aspose.Cells | Aspose.Cells table totals row example
// Developer Intent: Delete an existing totals row of a ListObject and recreate it with custom formulas for each numeric column using Aspose.Cells for .NET.
// Use Cases: Sales reports that need a double‑sum for quantity, average price, and maximum discount. | Financial worksheets where standard totals must be replaced with bespoke calculations. | Automated Excel exports that refresh custom aggregate values after data updates. | Data‑entry templates that enforce specific aggregate logic in the totals row.
// AI Prompts: Generate C# code with Aspose.Cells to remove a table's totals row, enable it again, and set custom formulas for Qty, Price, and Discount columns. | Explain how ShowTotals, TotalsCalculation, and SetCustomTotalsRowFormula work together in an Aspose.Cells ListObject. | Provide a step‑by‑step guide for adding a ListObject, toggling the totals row, and applying custom aggregate formulas in a .NET workbook.

using Aspose.Cells;
using Aspose.Cells.Tables;

// This example demonstrates how to create a workbook, add a ListObject covering A1:D5, hide the existing totals row, show it again, and assign custom formulas to the Qty, Price, and Discount columns using TotalsCalculation.Custom and SetCustomTotalsRowFormula. The workbook is then saved with the customized totals row.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate header row
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Qty");
        cells["C1"].PutValue("Price");
        cells["D1"].PutValue("Discount");

        // Populate sample numeric data
        string[] items = { "Apple", "Banana", "Cherry", "Date" };
        for (int i = 0; i < items.Length; i++)
        {
            int row = i + 1; // zero‑based index for data rows
            cells[row, 0].PutValue(items[i]);                 // Item name (text)
            cells[row, 1].PutValue((i + 1) * 10);             // Qty (numeric)
            cells[row, 2].PutValue((i + 1) * 2.5);            // Price (numeric)
            cells[row, 3].PutValue(0.05 * (i + 1));           // Discount (numeric)
        }

        // Add a table that covers the data range (A1:D5)
        int tableIndex = worksheet.ListObjects.Add(0, 0, items.Length, 3, true);
        ListObject table = worksheet.ListObjects[tableIndex];
        table.DisplayName = "SalesTable";

        // Show the totals row, then remove it
        table.ShowTotals = true;
        table.ShowTotals = false;   // Removal of the totals row

        // Re‑add the totals row
        table.ShowTotals = true;

        // ----- Set custom formulas for each numeric column -----

        // Column 1 (Qty) – custom total: double the sum of Qty
        ListColumn qtyColumn = table.ListColumns[1];
        qtyColumn.TotalsCalculation = TotalsCalculation.Custom;
        qtyColumn.SetCustomTotalsRowFormula("=SUM([Qty])*2", false, false);

        // Column 2 (Price) – custom total: average of Price
        ListColumn priceColumn = table.ListColumns[2];
        priceColumn.TotalsCalculation = TotalsCalculation.Custom;
        priceColumn.SetCustomTotalsRowFormula("=AVERAGE([Price])", false, false);

        // Column 3 (Discount) – custom total: maximum Discount
        ListColumn discountColumn = table.ListColumns[3];
        discountColumn.TotalsCalculation = TotalsCalculation.Custom;
        discountColumn.SetCustomTotalsRowFormula("=MAX([Discount])", false, false);

        // Save the workbook with the custom totals row
        workbook.Save("TableWithCustomTotals.xlsx");
    }
}
