// Title: Aspose.Cells C# – Enable Table Totals Row and Set a Distinct‑Count Formula
// Description: Creates a workbook, adds sample data, defines a ListObject, turns on the totals row, and uses SetCustomTotalsRowFormula with a SUMPRODUCT/COUNTIF expression to count unique values in the first column.
// Keywords: Aspose.Cells C# table totals row | custom totals formula | distinct count Excel formula | ListObject ShowTotals | SetCustomTotalsRowFormula
// Common Searches: Aspose.Cells show totals row C# | how to add distinct count to table totals in Aspose.Cells | custom totals calculation ListObject Aspose | C# Excel distinct count formula with Aspose.Cells
// Developer Intent: Add a ListObject, display its totals row, and configure a custom formula that returns the number of unique entries in a column.
// Use Cases: Generate summary sheets where the totals row reports unique category counts. | Automate Excel reports that need a distinct‑count metric without VBA. | Create data exports that include a quick‑look unique‑value statistic for analysts.
// AI Prompts: Write C# code using Aspose.Cells to add a table, enable its totals row, and apply a distinct‑count formula to a column. | Explain the interaction between TotalsCalculation.Custom and SetCustomTotalsRowFormula for unique value counting. | Suggest alternative Excel formulas for distinct counting that can be set with SetCustomTotalsRowFormula in Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, adds sample data, defines a ListObject, turns on the totals row, and uses SetCustomTotalsRowFormula with a SUMPRODUCT/COUNTIF expression to count unique values in the first column.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the table
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Value");
        string[] categories = { "A", "B", "A", "C", "B", "D" };
        for (int i = 0; i < categories.Length; i++)
        {
            cells[i + 1, 0].PutValue(categories[i]);   // Column A
            cells[i + 1, 1].PutValue(i + 10);          // Column B
        }

        // Add a table that includes the data range
        int tableIndex = sheet.ListObjects.Add(0, 0, categories.Length, 1, true);
        ListObject table = sheet.ListObjects[tableIndex];
        table.DisplayName = "MyTable";

        // Enable the totals row
        table.ShowTotals = true;

        // Configure a custom totals calculation for the first column (distinct count)
        ListColumn firstColumn = table.ListColumns[0];
        firstColumn.TotalsCalculation = TotalsCalculation.Custom;
        // Formula for distinct count using SUMPRODUCT and COUNTIF on the column reference
        firstColumn.SetCustomTotalsRowFormula("=SUMPRODUCT(1/COUNTIF([Category],[Category]))", false, false);

        // Save the workbook
        workbook.Save("TableDistinctCount.xlsx");
    }
}
