// Title: C# – Add a running‑total column to an Aspose.Cells ListObject with a previous‑row formula
// Description: This example creates a workbook, defines a table with ID and Amount columns, and inserts a dynamic formula into the third column that computes a running total by referencing the previous row. The first row copies the Amount value, subsequent rows add the current Amount to the prior RunningTotal, the formulas are evaluated, and the file is saved as RunningTotal.xlsx.
// Keywords: Aspose.Cells C# running total | PutCellFormula previous row | Excel ListObject cumulative sum | dynamic formula insertion Aspose | calculate running total Excel | smart markers table formula | C# Excel table formula reference | Aspose.Cells calculate formulas
// Common Searches: Aspose.Cells add running total column C# | Insert previous‑row formula in ListObject Aspose | C# cumulative sum in Excel table using Aspose.Cells | How to use PutCellFormula for running totals | Create running total column with Aspose.Cells
// Developer Intent: Insert a formula that references the previous row to produce a running‑total column in an Aspose.Cells table.
// Use Cases: Financial statements that display cumulative payments per period. | Invoice worksheets that automatically update a balance column as line items are added. | Sales dashboards showing progressive revenue totals across rows.
// AI Prompts: Show how to rewrite the formula using structured table references like [@Amount] instead of cell addresses. | Demonstrate adding the running‑total formula to an existing workbook without recreating the ListObject, handling pre‑populated data. | Explain how to apply the same running‑total logic to multiple tables in a workbook by iterating over each ListObject.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// This example creates a workbook, defines a table with ID and Amount columns, and inserts a dynamic formula into the third column that computes a running total by referencing the previous row. The first row copies the Amount value, subsequent rows add the current Amount to the prior RunningTotal, the formulas are evaluated, and the file is saved as RunningTotal.xlsx.
class RunningTotalDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Add headers for ID, Amount and RunningTotal
        ws.Cells["A1"].PutValue("ID");
        ws.Cells["B1"].PutValue("Amount");
        ws.Cells["C1"].PutValue("RunningTotal");

        // Sample data: {ID, Amount}
        int[,] data = { { 1, 10 }, { 2, 20 }, { 3, 15 }, { 4, 30 } };

        // Populate the worksheet with the sample data (starting at row 2)
        for (int i = 0; i < data.GetLength(0); i++)
        {
            ws.Cells[i + 1, 0].PutValue(data[i, 0]); // ID column (A)
            ws.Cells[i + 1, 1].PutValue(data[i, 1]); // Amount column (B)
        }

        // Create a ListObject (table) that covers the data range including headers
        // The range is A1:C5 because we have 4 data rows + 1 header row
        int tableIndex = ws.ListObjects.Add("A1", "C5", true);
        ListObject table = ws.ListObjects[tableIndex];

        // Insert a running‑total formula that references the previous row.
        // Row offset 0 = header row, so data rows start at offset 1.
        // Column offset 2 = third column (C) where the running total will appear.
        for (int rowOffset = 1; rowOffset <= data.GetLength(0); rowOffset++)
        {
            string formula;

            if (rowOffset == 1)
            {
                // First data row: running total equals the Amount value itself.
                // Excel row number = rowOffset + 1 (because Excel rows start at 1)
                formula = $"=B{rowOffset + 1}";
            }
            else
            {
                // Subsequent rows: Amount + previous RunningTotal.
                // Previous RunningTotal cell is in column C of the previous Excel row.
                formula = $"=B{rowOffset + 1}+C{rowOffset}";
            }

            // Apply the formula to the cell in column C of the current table row.
            table.PutCellFormula(rowOffset, 2, formula);
        }

        // Calculate all formulas so the running totals are materialized.
        wb.CalculateFormula();

        // Save the workbook.
        wb.Save("RunningTotal.xlsx");
    }
}
