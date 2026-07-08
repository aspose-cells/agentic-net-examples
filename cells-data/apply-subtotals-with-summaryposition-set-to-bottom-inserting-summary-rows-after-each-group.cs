using Aspose.Cells;
using System;

class SubtotalBottomDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add header row
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Item");
        cells["C1"].PutValue("Amount");

        // Sample data to be subtotaled
        object[,] data = new object[,]
        {
            { "A", "Item1", 100 },
            { "A", "Item2", 150 },
            { "B", "Item3", 200 },
            { "B", "Item4", 250 },
            { "A", "Item5", 120 }
        };

        // Populate the worksheet with the sample data (starting from row 2)
        for (int i = 0; i < data.GetLength(0); i++)
        {
            cells[i + 1, 0].PutValue(data[i, 0]); // Category
            cells[i + 1, 1].PutValue(data[i, 1]); // Item
            cells[i + 1, 2].PutValue(data[i, 2]); // Amount
        }

        // Define the cell area that includes the header and all data rows (A1:C6)
        CellArea area = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = data.GetLength(0), // 5 data rows + header = row index 5
            EndColumn = 2
        };

        // Apply subtotals:
        // - Group by the first column (Category) -> index 0
        // - Use SUM function on the third column (Amount) -> index 2
        // - Replace existing subtotals (true)
        // - Do not insert page breaks (false)
        // - Place summary rows below the data (true)
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, true, false, true);

        // Ensure the outline setting also places summary rows below the detail rows
        sheet.Outline.SummaryRowBelow = true;

        // Save the workbook to a file
        workbook.Save("SubtotalBottomDemo.xlsx");
    }
}