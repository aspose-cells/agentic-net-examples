// Title: Insert a Conditional SUMIF Formula per Row with Aspose.Cells in C# (Macro‑Style Routine)
// Description: A C# macro‑style example that creates a workbook, fills column A with values 1‑10 and column B with TRUE/FALSE flags, then writes an IF‑SUMIF formula into column C for every row. The formula returns the sum of all A‑column values where B is TRUE only when the current row's flag is TRUE; otherwise the cell stays blank. The routine calculates the formulas, prints results to the console, and saves the file as ConditionalSumIfDemo.xlsx.
// Keywords: Aspose.Cells C# conditional SUMIF | insert IF SUMIF formula Aspose.Cells | macro style routine Excel formulas .NET | set formula for each row Aspose.Cells | calculate formulas programmatically Aspose.Cells | dynamic Excel formulas C# | flag column conditional sum | Excel automation Aspose.Cells
// Common Searches: How to add an IF‑SUMIF formula to every row using Aspose.Cells C# | Aspose.Cells macro‑style routine for conditional formulas | Programmatically set and evaluate Excel formulas with Aspose.Cells | C# code to insert conditional SUMIF based on a flag column | Aspose.Cells example for dynamic row formulas
// Developer Intent: Programmatically add an IF‑SUMIF formula to each worksheet row that depends on a Boolean flag column, using Aspose.Cells for .NET.
// Use Cases: Create a summary column that shows the total of flagged values only when the current row is flagged. | Automate generation of workbooks where each row contains a dynamic formula referencing a data range, then evaluate the formulas before saving. | Produce Excel reports with pre‑calculated conditional totals for downstream analysis or BI tools.
// AI Prompts: Generate C# code with Aspose.Cells that writes an IF‑SUMIF formula to column C for each row, using column B as a TRUE/FALSE flag and column A as the source values. | Show how to trigger formula calculation in Aspose.Cells and retrieve the evaluated results for each row. | Explain how to modify the routine to detect the last data row automatically instead of using a hard‑coded count.

using System;
using Aspose.Cells;

namespace AsposeCellsMacroStyleRoutine
{
    // A C# macro‑style example that creates a workbook, fills column A with values 1‑10 and column B with TRUE/FALSE flags, then writes an IF‑SUMIF formula into column C for every row. The formula returns the sum of all A‑column values where B is TRUE only when the current row's flag is TRUE; otherwise the cell stays blank. The routine calculates the formulas, prints results to the console, and saves the file as ConditionalSumIfDemo.xlsx.
    public class InsertConditionalSumIf
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample data: Column A = values, Column B = flag (TRUE/FALSE)
                int totalRows = 10;
                for (int i = 0; i < totalRows; i++)
                {
                    // Value column (A)
                    cells[i, 0].PutValue(i + 1); // A1..A10 = 1..10

                    // Flag column (B) – TRUE for even rows, FALSE for odd rows
                    bool flag = (i % 2 == 0);
                    cells[i, 1].PutValue(flag);
                }

                // Determine the last data row (1‑based index for Excel formulas)
                int lastRowNumber = totalRows; // rows are 0‑based in API

                // Insert conditional SUMIF formula into Column C for each row
                for (int i = 0; i < totalRows; i++)
                {
                    // Build the formula string for the current row (Excel rows start at 1)
                    string formula = $"=IF($B{i + 1},SUMIF($B$2:$B${lastRowNumber + 1},TRUE,$A$2:$A${lastRowNumber + 1}),\"\")";

                    // Set the formula in column C (index 2)
                    cells[i, 2].Formula = formula;
                }

                // Calculate all formulas so that results are visible
                workbook.CalculateFormula();

                // Display results in console for verification
                Console.WriteLine("Row\tValue(A)\tFlag(B)\tResult(C)");
                for (int i = 0; i < totalRows; i++)
                {
                    Console.WriteLine($"{i + 1}\t{cells[i, 0].Value}\t{cells[i, 1].Value}\t{cells[i, 2].Value}");
                }

                // Save the workbook
                workbook.Save("ConditionalSumIfDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            InsertConditionalSumIf.Run();
        }
    }
}
