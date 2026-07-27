using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add header row
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Quantity");
        cells["C1"].PutValue("Price");

        // Add sample data rows
        string[,] data = {
            { "Apple",  "10", "0.5" },
            { "Banana", "20", "0.3" },
            { "Orange", "15", "0.4" }
        };

        for (int r = 0; r < data.GetLength(0); r++)
        {
            cells[r + 1, 0].PutValue(data[r, 0]);                     // Item
            cells[r + 1, 1].PutValue(int.Parse(data[r, 1]));         // Quantity (numeric)
            cells[r + 1, 2].PutValue(double.Parse(data[r, 2]));      // Price (numeric)
        }

        // Create a table that includes the header and data rows
        // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
        int tableIdx = sheet.ListObjects.Add(0, 0, data.GetLength(0), 2, true);
        ListObject table = sheet.ListObjects[tableIdx];
        table.DisplayName = "SalesTable";

        // Enable the totals row
        table.ShowTotals = true;

        // Configure sum calculations for numeric columns (Quantity and Price)
        table.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum; // Quantity column
        table.ListColumns[2].TotalsCalculation = TotalsCalculation.Sum; // Price column

        // Optionally set a label for the first column in the totals row
        table.ListColumns[0].TotalsRowLabel = "Grand Total";

        // Save the workbook
        workbook.Save("TotalsRowDemo.xlsx");
    }
}