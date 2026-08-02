using Aspose.Cells;
using System;

class SubtotalFlatListDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data (Category and Amount)
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Amount");
        object[,] data = new object[,]
        {
            { "A", 100 },
            { "A", 200 },
            { "B", 150 },
            { "B", 250 },
            { "C", 300 }
        };
        for (int i = 0; i < data.GetLength(0); i++)
        {
            cells[i + 1, 0].PutValue(data[i, 0]); // Category column
            cells[i + 1, 1].PutValue(data[i, 1]); // Amount column
        }

        // Define the cell area that contains the data (including header)
        CellArea area = CellArea.CreateCellArea("A1", "B5");

        // Apply subtotals:
        // - Group by the first column (Category) -> index 0
        // - Use SUM function
        // - Subtotal the second column (Amount) -> index 1
        // - replace: false (do not replace existing subtotals)
        // - pageBreaks: false (no page breaks between groups)
        // - summaryBelowData: false (summary rows placed above the group, producing a flat list)
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 1 }, false, false, false);

        // Ensure the outline does not treat the summary rows as hierarchical groups
        sheet.Outline.SummaryRowBelow = false;

        // Save the workbook to a file
        workbook.Save("SubtotalFlatList.xlsx");
    }
}