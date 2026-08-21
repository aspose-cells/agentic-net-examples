// Title: Assign row‑relative formula to column B in a loop and verify with Aspose.Cells for .NET
// Description: This example creates a workbook, fills A1‑A10 with numbers 1‑10, then iterates the rows to set the formula “=A{row}*2” in B1‑B10 using the Cell.Formula property. After calling Workbook.CalculateFormula it checks that each B cell equals its A counterpart multiplied by two, reports any mismatches, and saves the file as ColumnBFormulaLoop.xlsx.
// Keywords: Aspose.Cells | C# | Cell.Formula | loop | relative reference | calculate formulas | workbook save | formula validation | bulk assign formula
// Common Searches: Aspose.Cells set formula in a loop | apply same formula to a column programmatically | verify relative references in Aspose.Cells | C# bulk formula assignment | calculate and check formulas Aspose.Cells
// Developer Intent: Programmatically apply a row‑specific formula to a column and ensure the calculations are correct.
// Use Cases: Create a calculated column next to raw data | Generate per‑row totals in financial or inventory sheets | Automate QA of bulk‑assigned formulas before distribution | Populate derived values in reporting dashboards
// AI Prompts: Show C# code that loops through rows, assigns Cell.Formula referencing the same row in another column, runs Workbook.CalculateFormula, and flags mismatched results. | Provide a concise Aspose.Cells example for bulk‑setting a formula with relative references and validating the output. | Explain how to use Cell.Formula together with a loop to apply and verify row‑level calculations in a .NET workbook.

using System;
using Aspose.Cells;

// This example creates a workbook, fills A1‑A10 with numbers 1‑10, then iterates the rows to set the formula “=A{row}*2” in B1‑B10 using the Cell.Formula property. After calling Workbook.CalculateFormula it checks that each B cell equals its A counterpart multiplied by two, reports any mismatches, and saves the file as ColumnBFormulaLoop.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate column A with sample data (1 to 10)
        for (int row = 0; row < 10; row++)
        {
            cells[row, 0].PutValue(row + 1); // A1‑A10
        }

        // Loop through the same rows and assign a formula to column B.
        // The formula references the cell in column A of the same row.
        for (int row = 0; row < 10; row++)
        {
            // Using the Cell.Formula property (rule: Cell.Formula)
            cells[row, 1].Formula = $"=A{row + 1}*2"; // B1‑B10
        }

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Verify that each B cell contains the expected value (A * 2)
        bool allCorrect = true;
        for (int row = 0; row < 10; row++)
        {
            double aValue = Convert.ToDouble(cells[row, 0].Value);
            double bValue = Convert.ToDouble(cells[row, 1].Value);
            if (bValue != aValue * 2)
            {
                allCorrect = false;
                Console.WriteLine($"Mismatch at row {row + 1}: A={aValue}, B={bValue}");
            }
        }

        Console.WriteLine(allCorrect
            ? "All formulas evaluated correctly."
            : "There were mismatches in formula evaluation.");

        // Save the workbook
        workbook.Save("ColumnBFormulaLoop.xlsx");
    }
}
