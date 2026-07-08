using System;
using Aspose.Cells;

class SubtotalProductTopDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in columns I (group) and J (values)
        // Header row
        sheet.Cells["I1"].PutValue("Group");
        sheet.Cells["J1"].PutValue("Amount");

        // Sample data rows
        object[,] data = new object[,]
        {
            { "A", 2 },
            { "A", 3 },
            { "B", 4 },
            { "B", 5 },
            { "C", 6 }
        };

        // Fill the worksheet starting from row 2 (zero‑based index 1)
        for (int r = 0; r < data.GetLength(0); r++)
        {
            sheet.Cells[r + 1, 8].PutValue(data[r, 0]); // Column I (index 8)
            sheet.Cells[r + 1, 9].PutValue(data[r, 1]); // Column J (index 9)
        }

        // Define the cell area that contains the data (including headers)
        // Area starts at I1 (row 0, column 8) and ends at J6 (row 5, column 9)
        CellArea area = CellArea.CreateCellArea("I1", "J6");

        // Apply subtotals:
        // - Group by the first column of the area (Group column I) -> offset 0
        // - Use the Product function on the second column (Amount column J) -> offset 1
        // - Replace existing subtotals, no page breaks, place summary rows at the top (summaryBelowData = false)
        sheet.Cells.Subtotal(
            area,
            0,                                 // groupBy offset within the area
            ConsolidationFunction.Product,     // Product function
            new int[] { 1 },                   // totalList offset(s) within the area
            true,                              // replace existing subtotals
            false,                             // no page breaks between groups
            false);                            // summary rows placed above the group (top)

        // Save the workbook
        workbook.Save("SubtotalProductTopDemo.xlsx");
    }
}