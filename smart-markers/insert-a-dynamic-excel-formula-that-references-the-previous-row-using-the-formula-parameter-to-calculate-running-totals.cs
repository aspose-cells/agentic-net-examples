// Title: Programmatically insert a running‑total formula referencing the previous row in an Excel worksheet with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that adds a cumulative sum column where each cell formula adds the current value from column A to the total from the cell above in column B. | Show how to modify the example to start the running total at a specific row (e.g., row 5) and use mixed cell references in the inserted formulas. | Write a reusable method that takes a Worksheet, a source column index, and a target column index, then fills the target column with a running‑total formula that references the previous row.
// Common Searches: aspnet insert cumulative sum formula with Aspose.Cells C# | how to create running total column programmatically using Aspose.Cells for .NET | Aspose.Cells example adding previous row reference in Excel formula | C# generate dynamic Excel formula that adds previous row value using Aspose.Cells | calculate running totals in Excel workbook via Aspose.Cells API
// Tags: insert cumulative total formula Aspose.Cells | previous row reference in Excel formula .NET | generate cumulative totals programmatically C# | dynamic formula creation Aspose.Cells workbook | calculate cumulative sums using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// // This program creates a new workbook, fills column A with sample values, inserts a running‑total formula in column B that references the previous row (e.g., B2 = A2 + B1), recalculates all formulas, and saves the file as RunningTotal.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data placed in column A (zero‑based index 0)
            double[] values = { 10, 20, 15, 30, 25 };
            for (int i = 0; i < values.Length; i++)
            {
                cells[i, 0].PutValue(values[i]); // A column
            }

            // Insert running‑total formula in column B (zero‑based index 1)

            // First row: running total equals the first value
            cells[0, 1].Formula = "A1";

            // Subsequent rows: running total = previous total + current value
            // Build A1‑style formula for each row (e.g., B2 = A2 + B1)
            for (int row = 1; row < values.Length; row++)
            {
                string formula = $"A{row + 1}+B{row}";
                cells[row, 1].Formula = formula;
            }

            // Recalculate all formulas so the workbook contains the computed totals
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "RunningTotal.xlsx";

            // Save the workbook to a file (overwrite if exists)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
