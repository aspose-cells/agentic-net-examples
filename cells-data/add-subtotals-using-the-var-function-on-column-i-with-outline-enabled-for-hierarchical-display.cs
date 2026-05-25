using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // ----- Sample data (columns A to I) -----
        // Header row
        string[] headers = { "Category", "SubCategory", "Item", "Qty", "Price", "Discount", "Tax", "Region", "Value" };
        for (int col = 0; col < headers.Length; col++)
        {
            cells[0, col].PutValue(headers[col]);
        }

        // Data rows (10 rows)
        for (int row = 1; row <= 10; row++)
        {
            cells[row, 0].PutValue("Group" + ((row % 2) + 1));          // Category
            cells[row, 1].PutValue("Sub" + ((row % 3) + 1));            // SubCategory
            cells[row, 2].PutValue("Item" + row);                       // Item
            cells[row, 3].PutValue(row * 10);                           // Qty
            cells[row, 4].PutValue(row * 5);                            // Price
            cells[row, 5].PutValue(0);                                  // Discount
            cells[row, 6].PutValue(0);                                  // Tax
            cells[row, 7].PutValue("Region" + ((row % 2) + 1));         // Region
            cells[row, 8].PutValue(row * 100);                          // Value (Column I, zero‑based index 8)
        }

        // Define the cell area that includes the header and all data rows
        CellArea area = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = 10,
            EndColumn = 8
        };

        // Add subtotals:
        // - Group by the first column (Category) -> groupBy = 0
        // - Use the Var function (variance) on column I (index 8)
        // - Replace existing subtotals, insert page breaks, place summary below data
        worksheet.Cells.Subtotal(area, 0, ConsolidationFunction.Var, new int[] { 8 }, true, true, true);

        // Enable outline view (summary rows positioned below detail rows)
        worksheet.Outline.SummaryRowBelow = true;

        // Save the workbook
        workbook.Save("SubtotalVarOutline.xlsx");
    }
}