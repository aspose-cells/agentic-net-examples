// Title: Apply a shared‑style formula to the B1:D5 range and verify calculated values using Aspose.Cells for .NET
// AI Prompts: Assign the formula "=A{row}*10" to each cell in the B1:D5 block, recalculate the workbook, and list any cells where the computed value differs from the expected product. | Create a new Workbook, fill column A (A1:A5) with numbers 1‑5, set a relative multiplication formula for the entire B‑D range, run CalculateFormula, and programmatically assert that every cell equals its row index multiplied by 10. | After evaluating the formulas, save the workbook as SharedFormulaDemo.xlsx and output the absolute file path to the console.
// Common Searches: how to set the same formula for a range of cells using Aspose.Cells .NET | validate calculated values of a shared formula across multiple rows in Aspose.Cells | Aspose.Cells calculate formulas and compare results programmatically in C# | apply relative reference formula to B1:D5 with Aspose.Cells and verify output | save workbook after formula evaluation using Aspose.Cells for .NET
// Tags: apply identical formula to range Aspose.Cells | recalculate workbook formulas .NET | verify formula output C# | export workbook as xlsx Aspose.Cells | use relative cell reference in formula Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, populates column A with values 1‑5, assigns a relative formula "=A{row}*10" to every cell in the B1:D5 block, forces calculation, checks each result against the expected value, reports any mismatches, and saves the file as SharedFormulaDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column A with values 1..5 (rows 0‑4)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1); // A1=1, A2=2, ...
            }

            // Define the target range B1:D5 (rows 0‑4, columns 1‑3)
            // Apply a shared‑like formula: each cell = A(row)*10
            // Aspose.Cells .NET does not expose a Range.Formula property in some versions,
            // so we set the formula for each cell individually; Excel will adjust the relative row.
            for (int row = 0; row < 5; row++)
            {
                for (int col = 1; col <= 3; col++) // columns B, C, D
                {
                    // Use relative reference to the cell in column A of the same row
                    sheet.Cells[row, col].Formula = $"=A{row + 1}*10";
                }
            }

            // Force calculation of all formulas
            workbook.CalculateFormula();

            // Validate that every cell in B1:D5 contains the expected result
            bool allCorrect = true;
            for (int row = 0; row < 5; row++)
            {
                double expected = (row + 1) * 10; // A(row+1) * 10
                for (int col = 1; col <= 3; col++) // columns B, C, D
                {
                    double actual = sheet.Cells[row, col].DoubleValue;
                    if (Math.Abs(actual - expected) > 0.0001)
                    {
                        allCorrect = false;
                        Console.WriteLine($"Mismatch at {CellReference(row, col)}: expected {expected}, got {actual}");
                    }
                }
            }

            Console.WriteLine(allCorrect ? "All shared formula results are correct." : "There were mismatches.");

            // Save the workbook (lifecycle rule)
            string outputPath = "SharedFormulaDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper to convert zero‑based indices to Excel cell name (e.g., 0,1 -> B1)
    static string CellReference(int row, int column)
    {
        return CellsHelper.CellIndexToName(row, column);
    }
}
