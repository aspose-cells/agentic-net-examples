// Title: Logging evaluated formulas and their results in an Aspose.Cells workbook with C#
// Common Searches: C# how to print formula results after workbook.CalculateFormula in Aspose.Cells | Aspose.Cells iterate cells containing formulas and output values in .NET | retrieve and log each evaluated formula cell using Aspose.Cells C# example | display cell name, formula and calculated value after calculation with Aspose.Cells | log formula evaluation steps in an Excel file using Aspose.Cells for C#
// Tags: Aspose.Cells log evaluated formulas C# | iterate worksheet cells with formulas after CalculateFormula | retrieve calculated cell values Aspose.Cells | C# workbook formula evaluation output | display cell formula and result Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, writes numbers to A1 and A2, sets a SUM formula in A3, runs CalculateFormula, then iterates through all cells, printing each cell's address, its formula, and the evaluated result before saving the file as output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells with data
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);

                // Set a formula in A3
                sheet.Cells["A3"].Formula = "SUM(A1:A2)";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Log each formula cell and its calculated result
                LogFormulas(sheet);

                // Save the workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Logs formula cells and their evaluated values
        private static void LogFormulas(Worksheet sheet)
        {
            try
            {
                foreach (Cell cell in sheet.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        Console.WriteLine($"Evaluating cell {cell.Name} with formula: {cell.Formula}");
                        Console.WriteLine($"Result: {cell.Value}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Logging error: {ex.Message}");
            }
        }
    }
}
