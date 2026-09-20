// Title: Loop through rows in a workbook and assign a relative formula to column B with Aspose.Cells for .NET, then verify the calculated values
// AI Prompts: Generate C# code that iterates over rows 1‑10, sets each cell in column B to the formula "=A{row}*2" using Aspose.Cells, calls CalculateFormula, and checks that the resulting value equals the source cell multiplied by two. | Add error‑handling to an Aspose.Cells workbook that logs any mismatches between the expected A*2 result and the actual value in column B after formula evaluation.
// Common Searches: aspnet set same formula for entire column using Aspose.Cells loop | how to ensure relative cell references work when applying formulas with Aspose.Cells C# | validate calculated results after applying formulas in Aspose.Cells workbook | recalculate all formulas and compare expected values in Aspose.Cells .NET example
// Tags: loop assign formula Aspose.Cells C# | relative reference formula Aspose.Cells | validate calculated cell values Aspose.Cells | recalculate workbook formulas Aspose.Cells | save workbook after formula evaluation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Creates a new workbook, fills column A with numbers 1‑10, loops through rows to set each B cell to the formula "=A{row}*2", recalculates all formulas, verifies each B cell equals A*2, and saves the result as Result.xlsx.
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

                // Populate column A with sample data (1 to 10)
                for (int i = 0; i < 10; i++)
                {
                    cells[i, 0].PutValue(i + 1); // A1..A10
                }

                // Assign a formula to each cell in column B.
                // Using A1 notation: each B cell refers to the corresponding A cell.
                for (int row = 0; row < 10; row++)
                {
                    string formula = $"=A{row + 1}*2";
                    cells[row, 1].Formula = formula;
                }

                // Recalculate all formulas in the workbook
                workbook.CalculateFormula();

                // Verify that each cell in column B has the expected value (A*2)
                for (int row = 0; row < 10; row++)
                {
                    double expected = (row + 1) * 2;               // A value * 2
                    double actual = cells[row, 1].DoubleValue;    // B value after calculation

                    if (Math.Abs(expected - actual) > 0.0001)
                    {
                        Console.WriteLine($"Mismatch at B{row + 1}: expected {expected}, got {actual}");
                    }
                }

                // Define output file path
                string outputPath = "Result.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
