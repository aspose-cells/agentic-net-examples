// Title: Assign a German‑language SUMME formula to a cell using Cell.FormulaLocal in Aspose.Cells for .NET
// Common Searches: how to set a localized SUMME formula in Aspose.Cells C# | using FormulaLocal to apply German Excel functions with Aspose.Cells | Aspose.Cells calculate workbook after setting FormulaLocal property | C# example for culture‑specific formula assignment in an Excel file | set cell formula with localized function name in Aspose.Cells .NET
// Tags: Cell.FormulaLocal API Aspose.Cells | assigning localized formulas in Excel files | culture‑specific function names with Aspose.Cells | recalculating workbook after FormulaLocal assignment | saving workbook with localized formula

using System;
using Aspose.Cells;

// The example creates a new workbook, writes numeric values to cells A1‑A3, assigns the German localized formula "SUMME(A1:A3)" to cell A4 via the Cell.FormulaLocal property, recalculates all formulas, and saves the result as LocalizedFormula.xlsx.
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

            // Populate some cells with numeric values
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Set a formula using a localized function name (e.g., German "SUMME" for SUM)
            sheet.Cells["A4"].FormulaLocal = "SUMME(A1:A3)";

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook to a file
            workbook.Save("LocalizedFormula.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
