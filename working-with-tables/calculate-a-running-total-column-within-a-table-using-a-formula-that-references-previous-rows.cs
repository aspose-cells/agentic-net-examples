// Title: C# – Add a Running Total Column to an Aspose.Cells Table Using Structured References
// Description: The example creates a new workbook, builds a ListObject covering columns A‑B, fills column A with sample amounts, and defines column B as a calculated running‑total column using an IFERROR + OFFSET structured reference. Formulas are evaluated and the file is saved as RunningTotalTable.xlsx.
// Keywords: Aspose.Cells | C# | running total | structured reference | OFFSET function | ListObject | Excel table formula | cumulative sum | programmatic Excel | calculate formulas
// Common Searches: Aspose.Cells add running total column | C# OFFSET structured reference Aspose.Cells | calculate cumulative sum in Excel table programmatically | ListObject calculated column example | how to use IFERROR with OFFSET in Aspose.Cells
// Developer Intent: Create a calculated column in an Aspose.Cells ListObject that computes a running total based on the previous row’s value.
// Use Cases: Financial statements that show a cumulative balance for each transaction. | Inventory sheets tracking running stock levels as items are added or removed. | Sales dashboards displaying cumulative sales alongside daily figures.
// AI Prompts: Show how to modify the running‑total formula to skip rows where the Amount cell is blank. | Provide an alternative formula that uses SUM with a dynamic range instead of OFFSET. | Explain how to keep the running‑total column accurate when rows are inserted or deleted after the table is populated.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new workbook, builds a ListObject covering columns A‑B, fills column A with sample amounts, and defines column B as a calculated running‑total column using an IFERROR + OFFSET structured reference. Formulas are evaluated and the file is saved as RunningTotalTable.xlsx.
class RunningTotalExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Header row
        cells["A1"].PutValue("Amount");
        cells["B1"].PutValue("Running Total");

        // Populate sample amounts (1..9)
        for (int i = 2; i <= 10; i++)
        {
            cells[i - 1, 0].PutValue(i - 1); // Column A
        }

        // Create a table that includes both columns (A1:B10)
        // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
        int tableIndex = sheet.ListObjects.Add(0, 0, 9, 1, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Set the running total formula for the second column
        // Uses structured reference and OFFSET to refer to the previous row in the same column
        ListColumn runningTotalColumn = table.ListColumns[1];
        runningTotalColumn.Formula = "=IFERROR(OFFSET([@Running Total],-1,0),0)+[@Amount]";

        // Calculate all formulas so the running totals are materialized
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("RunningTotalTable.xlsx");
    }
}
