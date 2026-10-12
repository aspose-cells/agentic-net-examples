// Title: Add a cumulative running‑total column to an Excel table with Aspose.Cells for .NET using a SUM formula
// AI Prompts: Write C# code with Aspose.Cells that inserts a SUM($A$2:A{row}) formula in column B to compute a running total for each data row. | Show how to evaluate all formulas after adding the running‑total column and then save the workbook to a file. | Adapt the example to create an Aspose.Cells ListObject (Excel table) and calculate the running total within the table using structured references.
// Common Searches: aspnet calculate cumulative sum column in Excel with Aspose.Cells | c# Aspose.Cells insert running total formula that references previous rows | how to programmatically add a running total to an Excel table using Aspose.Cells | Aspose.Cells evaluate formulas after adding cumulative total column
// Tags: cumulative running total formula Aspose.Cells | insert SUM($A$2:Arow) column C# | evaluate formulas workbook Aspose.Cells | create ListObject running total Aspose.Cells | populate column A sample data Aspose.Cells

using Aspose.Cells;
using System;

// // Demonstrates creating a new workbook, filling column A with sample numbers, adding a cumulative SUM formula in column B for each row to produce a running total, calculating all formulas, and saving the workbook as RunningTotal.xlsx.
class RunningTotalExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add headers for the source values and the running total column
        cells["A1"].PutValue("Value");
        cells["B1"].PutValue("RunningTotal");

        // Sample data to demonstrate the running total
        double[] values = { 10, 20, 15, 30, 25 };

        // Populate column A with the sample values (starting at row 2)
        for (int i = 0; i < values.Length; i++)
        {
            // Row index = i + 2 (because Excel rows are 1‑based and row 1 is header)
            cells[i + 2, 0].PutValue(values[i]); // Column 0 = A
        }

        // Determine the last row that contains data (including header row)
        int lastDataRow = values.Length + 1; // +1 because row 1 is the header

        // Insert a running total formula in column B for each data row
        // Formula used: SUM($A$2:A{currentRow})
        // This sums all values from A2 down to the current row, giving a cumulative total.
        for (int row = 2; row <= lastDataRow; row++)
        {
            // Column 1 = B
            cells[row - 1, 1].Formula = $"SUM($A$2:A{row})";
        }

        // Evaluate all formulas so the workbook contains the calculated results
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("RunningTotal.xlsx");
    }
}
