// Title: How to assign a German‑language SUM formula using Cell.FormulaLocal in Aspose.Cells for .NET
// Common Searches: c# asp.net set German SUMME function in Excel workbook using Aspose.Cells | how to use Aspose.Cells FormulaLocal to write culture‑specific formulas | apply localized Excel formulas and recalculate workbook with Aspose.Cells .NET | save Excel file with localized function names using Aspose.Cells API | calculate workbook after assigning FormulaLocal formula in C#
// Tags: assign localized SUM formula Aspose.Cells | culture‑specific formula API Aspose.Cells | evaluate formulas after localized assignment Aspose.Cells | export Excel with culture‑specific formulas Aspose.Cells | handling culture‑specific function names in .NET Excel

using Aspose.Cells;
using System;

// The example creates a new workbook, writes numeric values to cells A1 and A2, sets a German‑language SUM formula (SUMME) in cell A3 via the Cell.FormulaLocal property, recalculates all formulas, and saves the workbook as LocalizedFormula.xlsx.
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

            // Set a formula using a localized function name.
            // Example: German function name for SUM is "SUMME"
            sheet.Cells["A3"].FormulaLocal = "SUMME(A1:A2)";

            // Recalculate the workbook to evaluate the formula
            workbook.CalculateFormula();

            // Save the workbook to a file
            workbook.Save("LocalizedFormula.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
