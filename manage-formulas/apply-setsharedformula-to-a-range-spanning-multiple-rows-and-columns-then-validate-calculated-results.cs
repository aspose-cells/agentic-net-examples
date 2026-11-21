// Title: Apply a shared formula to a multi‑row, multi‑column range and verify results with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shared formula for B2:D5 range | how to validate calculated values after workbook.CalculateFormula in Aspose.Cells | assign same formula to multiple rows and columns using Aspose.Cells API | C# example for using shared formulas with relative row references in Aspose.Cells | check formula evaluation results programmatically with Aspose.Cells
// Tags: shared formula assignment Aspose.Cells C# | apply formula to range Aspose.Cells | calculate workbook formulas Aspose.Cells | validate cell calculation results C# | relative row reference formula Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a Workbook, fills A2:A5 with incremental numbers, assigns the formula "=A{row}*2" to every cell in the B2:D5 block using shared formulas, recalculates all formulas with workbook.CalculateFormula(), and then iterates through the range to confirm each cell's value equals the corresponding A‑column value multiplied by two, reporting any mismatches.
class SetSharedFormulaExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column A with base values (A2:A5).
            for (int row = 1; row <= 4; row++) // rows are zero‑based; row 1 = A2
            {
                cells[row, 0].PutValue(row * 10); // A2=10, A3=20, A4=30, A5=40
            }

            // Define the target range B2:D5 (rows 1‑4, columns 1‑3).
            int firstRow = 1, firstCol = 1;
            int totalRows = 4, totalCols = 3;

            // Apply a formula to each cell in the range.
            // The formula refers to the cell in column A of the same row: =A2*2
            for (int r = 0; r < totalRows; r++)
            {
                int worksheetRow = firstRow + r;
                for (int c = 0; c < totalCols; c++)
                {
                    int worksheetCol = firstCol + c;
                    // Build a formula that points to column A of the current row.
                    string formula = $"=A{worksheetRow + 1}*2";
                    cells[worksheetRow, worksheetCol].Formula = formula;
                }
            }

            // Recalculate all formulas in the workbook.
            workbook.CalculateFormula();

            // Validate the calculated results.
            bool allCorrect = true;
            for (int r = 0; r < totalRows; r++)
            {
                int worksheetRow = firstRow + r;
                double baseValue = cells[worksheetRow, 0].DoubleValue;
                double expected = baseValue * 2;

                for (int c = 0; c < totalCols; c++)
                {
                    int worksheetCol = firstCol + c;
                    double actual = cells[worksheetRow, worksheetCol].DoubleValue;

                    if (Math.Abs(actual - expected) > 1e-9)
                    {
                        allCorrect = false;
                        Console.WriteLine($"Mismatch at {cells[worksheetRow, worksheetCol].Name}: expected {expected}, got {actual}");
                    }
                }
            }

            if (allCorrect)
            {
                Console.WriteLine("All shared formula calculations are correct.");
            }

            // Optional: Save the workbook to verify manually.
            // workbook.Save("SharedFormulaResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
