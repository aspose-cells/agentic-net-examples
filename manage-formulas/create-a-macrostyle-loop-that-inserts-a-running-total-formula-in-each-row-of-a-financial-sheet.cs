// Title: Add a cumulative running‑total formula to each row of a financial worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that iterates over all data rows and writes a `=SUM($C$2:C{row})` formula into column D, then forces calculation. | Show how to modify the loop so the running total can start from any given row and column while keeping the workbook auto‑calculated. | Generate a reusable method that accepts a worksheet, start row, amount column index, and total column index, and inserts cumulative sum formulas using Aspose.Cells.
// Common Searches: aspnet c# how to programmatically add a cumulative sum column with Aspose.Cells | Aspose.Cells loop to insert running total formula in Excel sheet | set SUM($C$2:Crow) formula for each row using Aspose.Cells .NET | auto calculate workbook after inserting formulas Aspose.Cells C# | change start row for running total column in Aspose.Cells example
// Tags: insert cumulative sum formula Aspose.Cells | loop to set SUM formula column D C# | auto calculate workbook Aspose.Cells | financial worksheet running total automation | parameterized running total method Aspose.Cells

using System;
using Aspose.Cells;

namespace FinancialRunningTotal
{
    // Demonstrates creating a workbook, populating column C with sample amounts, looping through each data row to place a cumulative `=SUM($C$2:Crow)` formula in column D, triggering calculation, and saving the file as FinancialRunningTotal.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a worksheet named "Financial"
            Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Add()];
            sheet.Name = "Financial";

            // Example data: populate column C with some amounts starting from row 2
            // (In a real scenario the data would already exist)
            for (int i = 2; i <= 10; i++)
            {
                sheet.Cells[i, 2].PutValue(100 * i); // Column C (index 2)
            }

            // Insert running total formula in column D for each row
            // Formula: =SUM($C$2:C2) -> cumulative sum up to the current row
            for (int row = 2; row <= sheet.Cells.MaxDataRow; row++)
            {
                // Build the formula string with absolute reference to the start row
                string formula = $"=SUM($C$2:C{row})";

                // Place the formula in column D (index 3) of the current row
                sheet.Cells[row, 3].Formula = formula;
            }

            // Auto-calculate the workbook to evaluate the formulas
            workbook.CalculateFormula();

            // Save the workbook to a file
            workbook.Save("FinancialRunningTotal.xlsx");
        }
    }
}
