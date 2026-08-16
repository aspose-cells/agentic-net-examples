// Title: C# Aspose.Cells example: Loop column A and assign row‑based multiplication formulas to column B
// Description: Creates a workbook, fills A1‑A10 with 1‑10, loops each row to set B cells with =A{row}*{row}, calculates all formulas, and saves the file as MultiplicationResults.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | set cell formula Aspose.Cells | loop rows Excel C# | row based multiplication formula | calculate formulas Aspose.Cells | save workbook Aspose.Cells | Excel automation .NET | dynamic formula assignment
// Common Searches: Aspose.Cells set formula in C# loop rows | how to apply multiplication formula per row Aspose.Cells | calculate and save Excel workbook with formulas C# | loop through column A and write formulas to column B Aspose
// Developer Intent: Generate a workbook, populate column A, apply a row‑specific multiplication formula to column B, evaluate the formulas, and persist the result.
// Use Cases: Create a numeric list and automatically compute each value's square in the adjacent column. | Apply a custom calculation that incorporates the current row number for per‑row analytics. | Automate Excel reporting where formulas are generated programmatically and saved without manual editing.
// AI Prompts: Write C# code with Aspose.Cells that fills column A with numbers 1‑N, sets column B formulas as =Arow*row, calculates all formulas, and saves the workbook. | Generate a reusable method that receives a worksheet and target column index, iterates over used rows, and assigns a row‑based multiplication formula using the Formula property. | Explain the differences between the Formula property and the SetFormula method when assigning formulas in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, fills A1‑A10 with 1‑10, loops each row to set B cells with =A{row}*{row}, calculates all formulas, and saves the file as MultiplicationResults.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column A with sample values (1, 2, 3, ...)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(i + 1); // A1..A10
            }

            // Loop through each used row in column A
            // For each row, set a multiplication formula in column B:
            //   =A{rowNumber}*{rowNumber}
            for (int row = 0; row < 10; row++)
            {
                string formula = $"=A{row + 1}*{row + 1}";
                // Set the formula; using the Formula property avoids null options issue
                cells[row, 1].Formula = formula;
            }

            // Calculate all formulas so that column B contains the results
            workbook.CalculateFormula();

            // Save the workbook with the computed results (lifecycle: save)
            workbook.Save("MultiplicationResults.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
