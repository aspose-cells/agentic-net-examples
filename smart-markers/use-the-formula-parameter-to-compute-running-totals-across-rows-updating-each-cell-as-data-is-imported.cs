// Title: Import an integer array into Excel and generate a running‑total column with per‑row formulas using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an integer array into column A of a new workbook and assigns a cumulative‑sum formula to each cell in column B using Aspose.Cells. | Show how to build a formula string that references the previous row’s total (e.g., =B{row‑1}+A{row}) and set it on the target cell with Aspose.Cells. | Demonstrate calculating all formulas, printing the values, and saving the workbook as an .xlsx file with Aspose.Cells.
// Common Searches: how to add a running total column in Excel with Aspose.Cells C# | Aspose.Cells set formula for each row programmatically | cumulative sum calculation while importing data using Aspose.Cells .NET | example of dynamic Excel formulas based on previous row with Aspose.Cells | save workbook after evaluating formulas Aspose.Cells C#
// Tags: Aspose.Cells running total formula | Aspose.Cells assign per‑row formula | Aspose.Cells import integer array | Aspose.Cells calculate cumulative sum | Aspose.Cells save workbook with formulas

using System;
using Aspose.Cells;

namespace RunningTotalExample
{
    // The example creates a new workbook, imports an integer array into column A, writes a running‑total formula in column B for each row (adding the current value to the previous total), evaluates all formulas, prints the results to the console, and saves the file as RunningTotal.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data to be imported (could come from any source)
            int[] values = { 10, 20, 15, 30, 25 };

            // Import data row by row and set a running‑total formula in the adjacent column
            for (int i = 0; i < values.Length; i++)
            {
                // Excel rows are 1‑based, but Cells uses 0‑based indexes.
                // Column A (index 0) will hold the raw values.
                cells[i, 0].PutValue(values[i]);

                // Column B (index 1) will hold the running total.
                if (i == 0)
                {
                    // First row: running total equals the first value.
                    cells[i, 1].PutValue(values[i]);
                }
                else
                {
                    // For subsequent rows set a formula:
                    //   =B{previousRow}+A{currentRow}
                    //   where B{previousRow} is the running total of the row above.
                    int currentExcelRow = i + 1;          // e.g., i=1 => row 2
                    int previousExcelRow = i;            // e.g., i=1 => row 1
                    string formula = $"=B{previousExcelRow}+A{currentExcelRow}";
                    // Assign the formula to the cell.
                    cells[i, 1].Formula = formula;
                }
            }

            // Calculate all formulas so that the running totals are materialized.
            workbook.CalculateFormula();

            // Optional: display the results in the console.
            Console.WriteLine("Row\tValue\tRunning Total");
            for (int i = 0; i < values.Length; i++)
            {
                Console.WriteLine($"{i + 1}\t{cells[i, 0].IntValue}\t{cells[i, 1].IntValue}");
            }

            // Save the workbook (uses the standard create‑save lifecycle).
            workbook.Save("RunningTotal.xlsx");
        }
    }
}
