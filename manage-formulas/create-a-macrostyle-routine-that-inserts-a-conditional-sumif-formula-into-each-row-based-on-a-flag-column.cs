// Title: C# Aspose.Cells macro‑style routine to insert a conditional SUMIF in each row
// Description: Creates a new workbook, fills column A with numbers and column B with Boolean flags, builds absolute A and B ranges, and loops through every row to place an IF‑wrapped SUMIF formula in column C that sums column A only when the flag in column B is TRUE. The routine then calculates all formulas, writes results to the console, and saves the file as ConditionalSumIfDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | conditional SUMIF | macro style routine | insert formula programmatically | IF SUMIF Excel | range strings Aspose | calculate formulas | flag column aggregation
// Common Searches: How to add a conditional SUMIF formula to each row with Aspose.Cells C# | Aspose.Cells macro‑style code to insert IF‑SUMIF based on a flag column | Programmatically create absolute range strings for SUMIF in Aspose.Cells | Calculate and retrieve results of inserted formulas using Aspose.Cells .NET | Insert row‑wise conditional aggregation formula in Excel via C#
// Developer Intent: Programmatically place an IF‑wrapped SUMIF formula in column C for every row, using a Boolean flag in column B, with Aspose.Cells.
// Use Cases: Generate a dynamic total column that adds values from column A only when a corresponding flag in column B is TRUE. | Apply the same conditional aggregation logic to any worksheet without manual Excel editing. | Export a workbook that already contains calculated results for downstream reporting or analytics.
// AI Prompts: Write a C# method using Aspose.Cells that inserts an IF‑SUMIF formula into each row based on a flag column and then calculates the workbook. | Adapt the routine to accept a custom start row and allow the flag criteria to be a user‑defined string instead of the hard‑coded TRUE. | Explain how to modify the absolute range strings when the data begins at a different row or includes header rows.

using System;
using Aspose.Cells;

namespace AsposeCellsMacroStyleRoutine
{
    // Creates a new workbook, fills column A with numbers and column B with Boolean flags, builds absolute A and B ranges, and loops through every row to place an IF‑wrapped SUMIF formula in column C that sums column A only when the flag in column B is TRUE. The routine then calculates all formulas, writes results to the console, and saves the file as ConditionalSumIfDemo.xlsx.
    public class ConditionalSumIfRoutine
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data (10 rows)
                for (int row = 0; row < 10; row++)
                {
                    // Column A: numeric value
                    cells[row, 0].PutValue(10 * (row + 1));

                    // Column B: flag (TRUE for even rows, FALSE for odd rows)
                    bool flag = (row % 2 == 0);
                    cells[row, 1].PutValue(flag);
                }

                // Determine the last data row (zero‑based index)
                int lastRow = cells.MaxDataRow;

                // Build absolute ranges for SUMIF (Excel rows are 1‑based)
                string aRange = $"$A$2:$A${lastRow + 1}";
                string bRange = $"$B$2:$B${lastRow + 1}";

                // Insert conditional SUMIF formula into column C for each row
                for (int row = 0; row <= lastRow; row++)
                {
                    int excelRow = row + 1; // Excel row number (1‑based)
                    string conditionCell = $"$B${excelRow}";
                    string formula = $"=IF({conditionCell}=TRUE, SUMIF({bRange},TRUE,{aRange}), 0)";
                    cells[row, 2].Formula = formula;
                }

                // Calculate all formulas
                workbook.CalculateFormula();

                // Output results to console for verification
                Console.WriteLine("Row\tValue(A)\tFlag(B)\tConditionalSum(C)");
                for (int row = 0; row <= lastRow; row++)
                {
                    Console.WriteLine($"{row + 1}\t{cells[row, 0].Value}\t{cells[row, 1].Value}\t{cells[row, 2].Value}");
                }

                // Save the workbook
                workbook.Save("ConditionalSumIfDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ConditionalSumIfRoutine.Run();
        }
    }
}
