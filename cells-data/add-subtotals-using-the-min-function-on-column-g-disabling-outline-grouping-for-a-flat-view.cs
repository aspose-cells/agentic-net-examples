// Title: Create a flat‑view Excel file with minimum subtotals on column G grouped by column F using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that groups rows by column F and inserts a Min subtotal for column G while keeping the worksheet outline flat. | Show how to turn off outline grouping after adding subtotals so the summary rows stay inline in the same sheet using Aspose.Cells .NET.
// Common Searches: how to add a min subtotal for a column in Aspose.Cells C# | Aspose.Cells disable outline grouping for flat view after subtotal | C# Aspose.Cells subtotal by group column without inserting extra rows
// Tags: min subtotal calculation Aspose.Cells C# | column grouping subtotal Aspose.Cells | outline grouping off for flat view Aspose.Cells | subtotal without extra summary rows Aspose.Cells

using Aspose.Cells;

// The example creates a new workbook, fills columns F (group) and G (values) with sample data, defines a cell area, adds subtotals that compute the minimum of column G for each group in column F, disables outline grouping to produce a flat view, and saves the file as SubtotalMinFlatView.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // -------------------------------------------------
        // Sample data preparation (columns F = Group, G = Value)
        // -------------------------------------------------
        cells["F1"].PutValue("Group");   // Header for grouping column (column G in zero‑based index is 5)
        cells["G1"].PutValue("Value");   // Header for values column (column G in zero‑based index is 6)

        // Populate 10 rows of sample data
        for (int i = 0; i < 10; i++)
        {
            // Alternate groups "A" and "B" in column F
            cells[i + 1, 5].PutValue(i % 2 == 0 ? "A" : "B");
            // Some numeric values in column G
            cells[i + 1, 6].PutValue(100 + i * 10);
        }

        // -------------------------------------------------
        // Define the range that contains the data (including headers)
        // StartRow = 0, StartColumn = 5 (F), EndRow = 10, EndColumn = 6 (G)
        // -------------------------------------------------
        CellArea area = CellArea.CreateCellArea(0, 5, 10, 6);

        // -------------------------------------------------
        // Add subtotals:
        //   - Group by the first column of the area (column F)
        //   - Use the Min function on the second column of the area (column G)
        //   - Do not replace existing subtotals, no page breaks, and do not place summary rows below data
        // -------------------------------------------------
        cells.Subtotal(
            area,
            0,                                 // groupBy: first column of the area (column F)
            ConsolidationFunction.Min,         // function: Min
            new int[] { 1 },                   // totalList: second column of the area (column G)
            false,                             // replace existing subtotals
            false,                             // add page breaks between groups
            false                              // place summary rows below data
        );

        // -------------------------------------------------
        // Disable outline grouping for a flat view
        // Setting SummaryRowBelow to false prevents the outline from inserting extra rows
        // -------------------------------------------------
        sheet.Outline.SummaryRowBelow = false;

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("SubtotalMinFlatView.xlsx");
    }
}
