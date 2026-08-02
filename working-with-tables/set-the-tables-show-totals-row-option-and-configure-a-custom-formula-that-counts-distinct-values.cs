// Title: C# – Show Table Totals Row and Apply a Custom Distinct‑Count Formula with Aspose.Cells
// Description: Creates a workbook, inserts sample data, defines a ListObject (table), enables its totals row, sets the first column to use a custom totals calculation, and applies a SUMPRODUCT/COUNTIF formula that returns the number of unique values in the Category column before saving the file.
// Keywords: Aspose.Cells C# table totals row | custom totals calculation | distinct count formula | ListObject ShowTotals | SetCustomTotalsRowFormula | SUMPRODUCT COUNTIF Excel | Aspose.Cells example
// Common Searches: how to enable totals row in Aspose.Cells table | set custom distinct count formula C# Aspose.Cells | ListObject ShowTotals property example | apply SUMPRODUCT COUNTIF with Aspose.Cells | custom totals calculation for Excel table using Aspose
// Developer Intent: Add a totals row to a ListObject and configure a custom formula that returns the count of unique entries in a specified column.
// Use Cases: Generate a summary row that shows the number of different categories in a data set. | Build automated Excel reports where the totals line provides a distinct‑count metric for any categorical field. | Create dynamic workbooks that adapt the distinct‑count calculation as rows are added or removed.
// AI Prompts: Write C# code with Aspose.Cells to create a table, turn on its totals row, and set a custom distinct‑count formula for a column. | Explain the parameters of SetCustomTotalsRowFormula and how they affect A1‑style and culture‑independent formulas. | Suggest a locale‑aware distinct‑count expression for Excel and demonstrate its implementation with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, inserts sample data, defines a ListObject (table), enables its totals row, sets the first column to use a custom totals calculation, and applies a SUMPRODUCT/COUNTIF formula that returns the number of unique values in the Category column before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data for the table
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Value");

        string[] categories = { "A", "B", "A", "C", "B", "A" };
        int[] values = { 10, 20, 30, 40, 50, 60 };

        for (int i = 0; i < categories.Length; i++)
        {
            cells[i + 1, 0].PutValue(categories[i]); // Column A
            cells[i + 1, 1].PutValue(values[i]);    // Column B
        }

        // Add a table that includes the data range
        // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
        int tableIndex = worksheet.ListObjects.Add(0, 0, categories.Length, 1, true);
        ListObject table = worksheet.ListObjects[tableIndex];
        table.DisplayName = "MyTable";

        // Enable the totals row for the table
        table.ShowTotals = true;

        // Configure a custom totals calculation for the first column (Category)
        ListColumn categoryColumn = table.ListColumns[0];
        categoryColumn.TotalsCalculation = TotalsCalculation.Custom;

        // Set a custom formula that counts distinct values in the Category column
        // The formula uses SUMPRODUCT with COUNTIF to achieve distinct count
        // isR1C1 = false (A1 style), isLocal = false (invariant culture)
        categoryColumn.SetCustomTotalsRowFormula("=SUMPRODUCT(1/COUNTIF([Category],[Category]))", false, false);

        // Save the workbook to a file
        workbook.Save("TableDistinctCount.xlsx");
    }
}
