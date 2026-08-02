using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data for the table
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Quantity");
        cells["C1"].PutValue("Price");

        for (int row = 2; row <= 6; row++)
        {
            cells[$"A{row}"].PutValue($"Item{row - 1}");
            cells[$"B{row}"].PutValue(row * 10);          // Quantity values
            cells[$"C{row}"].PutValue(row * 2.5);         // Price values
        }

        // Add a ListObject (table) that spans the data range
        int tableIndex = worksheet.ListObjects.Add("A1", "C6", true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Enable the totals row for the table
        table.ShowTotals = true;

        // Select the column for which we want a custom average in the totals row (Quantity column)
        ListColumn quantityColumn = table.ListColumns[1]; // Index 1 corresponds to the second column ("Quantity")

        // Set the totals calculation type to Custom
        quantityColumn.TotalsCalculation = TotalsCalculation.Custom;

        // Define a custom formula that calculates the average of the "Quantity" column
        // The formula uses the column name in brackets as required by Aspose.Cells table formulas
        quantityColumn.SetCustomTotalsRowFormula("=AVERAGE([Quantity])", false, false);

        // Save the workbook to a file
        workbook.Save("CustomTotalsAverage.xlsx");
    }
}