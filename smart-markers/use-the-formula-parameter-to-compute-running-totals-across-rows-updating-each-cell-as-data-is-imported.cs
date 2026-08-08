// Title: C# – Compute Running Totals with PutCellFormula in an Aspose.Cells Table
// Description: Creates a workbook, adds a header and numeric data, defines a ListObject covering columns A‑B, and uses PutCellFormula to insert a cumulative SUM formula into each row of the RunningTotal column. The workbook then calculates all formulas and saves the file as RunningTotal.xlsx.
// Keywords: Aspose.Cells | C# | PutCellFormula | running total | cumulative sum | ListObject | Excel table formula | CalculateFormula | smart markers
// Common Searches: Aspose.Cells putcellformula running total example | C# cumulative sum column Aspose.Cells | how to add running total to Excel table with Aspose.Cells | calculate progressive total in Aspose.Cells .NET | smart markers running total formula
// Developer Intent: Insert a per‑row formula that calculates a running total and evaluate it using Aspose.Cells.
// Use Cases: Import a series of numbers and automatically generate a column that shows the cumulative sum for each row. | Build an Excel table where the RunningTotal column updates instantly when source values change, using PutCellFormula and CalculateFormula. | Create a financial statement that displays the progressive total of transactions without manually writing formulas.
// AI Prompts: Generate C# code that uses structured table references (e.g., [@Value]) instead of absolute A2 references for the running total formula. | Show how to add a column that calculates the percentage of each running total relative to the final total. | Provide a version that automatically adjusts running‑total formulas when rows are inserted or deleted.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, adds a header and numeric data, defines a ListObject covering columns A‑B, and uses PutCellFormula to insert a cumulative SUM formula into each row of the RunningTotal column. The workbook then calculates all formulas and saves the file as RunningTotal.xlsx.
class RunningTotalExample
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Header row
        cells[0, 0].PutValue("Value");
        cells[0, 1].PutValue("RunningTotal");

        // Sample data to be imported (10 rows)
        int rowCount = 10;
        for (int i = 0; i < rowCount; i++)
        {
            // Import data into the first column
            cells[i + 1, 0].PutValue((i + 1) * 10);
        }

        // Create a table that includes the header and data (range A1:B{rowCount+1})
        int tableIndex = sheet.ListObjects.Add(0, 0, rowCount, 1, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // For each data row, set a formula in the RunningTotal column that computes the running sum
        // Using PutCellFormula(rowOffset, columnOffset, formula) rule
        for (int i = 0; i < rowCount; i++)
        {
            // Row offset is i (0‑based within the table, excluding header)
            // Column offset 1 corresponds to the second column (RunningTotal)
            // Formula: =SUM($A$2:A{currentRow})
            // CurrentRow in worksheet coordinates = i + 2 (because row 1 is header, row 2 is first data)
            string formula = $"=SUM($A$2:A{i + 2})";
            table.PutCellFormula(i, 1, formula);
        }

        // Calculate all formulas so that the running totals are materialized
        workbook.CalculateFormula();

        // Save the workbook (lifecycle rule: save)
        workbook.Save("RunningTotal.xlsx");
    }
}
