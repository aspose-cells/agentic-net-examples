// Title: Apply a shared formula to a multi‑row, multi‑column range and verify results with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# SetSharedFormula for a block of cells | How to validate calculated values after applying a shared formula in .NET | Recalculate workbook and check shared formula results using Aspose.Cells | Apply the same formula to columns C and D across rows 1 to 5 in C# | Save an Excel file after using SetSharedFormula with Aspose.Cells
// Tags: shared formula SetSharedFormula Aspose.Cells | apply shared formula to range C1:D5 | recalculate workbook formulas Aspose.Cells | check shared formula output Aspose.Cells C# | export workbook to .xlsx Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, fills columns A and B with sample data, uses SetSharedFormula to apply the formula "A1+B1" across the C1:D5 range, recalculates the workbook, validates each result against the expected sum, reports any mismatches, and finally saves the file as SharedFormulaDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Fill sample data in columns A and B (rows 1‑5)
            for (int i = 0; i < 5; i++)
            {
                ws.Cells[i, 0].PutValue(i + 1);          // Column A
                ws.Cells[i, 1].PutValue((i + 1) * 10);   // Column B
            }

            // Apply a shared formula =A1+B1 to range C1:D5
            // The shared formula is defined relative to the top‑left cell (A1)
            int firstRow = 0;      // row index of A1
            int firstColumn = 0;   // column index of A1
            for (int row = 0; row < 5; row++)
            {
                for (int col = 2; col <= 3; col++) // Columns C and D
                {
                    ws.Cells[row, col].SetSharedFormula("A1+B1", firstRow, firstColumn);
                }
            }

            // Recalculate workbook to evaluate formulas
            wb.CalculateFormula();

            // Validate calculated results
            bool allOk = true;
            for (int row = 0; row < 5; row++)
            {
                double expected = ws.Cells[row, 0].DoubleValue + ws.Cells[row, 1].DoubleValue;
                for (int col = 2; col <= 3; col++) // Columns C and D
                {
                    double actual = ws.Cells[row, col].DoubleValue;
                    if (Math.Abs(actual - expected) > 1e-9)
                    {
                        allOk = false;
                        Console.WriteLine($"Mismatch at {ws.Cells[row, col].Name}: expected {expected}, got {actual}");
                    }
                }
            }

            Console.WriteLine(allOk ? "All shared formula results are correct." : "There were mismatches.");

            // Save the workbook (demonstrating lifecycle)
            string outputPath = "SharedFormulaDemo.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
