// Title: Set a Portuguese-language SUM formula in an Excel workbook using Cell.FormulaLocal with Aspose.Cells for .NET
// AI Prompts: Insert a Portuguese SUM (SOMA) formula into cell A1 via Cell.FormulaLocal, then recalculate the workbook. | Configure the workbook's Settings.CultureInfo to pt-PT, apply a semicolon‑separated localized formula, and save the file. | Demonstrate how to evaluate culture‑specific formulas in Aspose.Cells by using FormulaLocal and Workbook.CalculateFormula.
// Common Searches: how to assign a Portuguese SUM formula using Aspose.Cells C# | using Cell.FormulaLocal with pt-PT culture in Aspose.Cells | example of semicolon separated formulas in Aspose.Cells .NET | set workbook culture for localized functions in Aspose.Cells | calculate localized Excel formulas programmatically with Aspose.Cells
// Tags: Cell.FormulaLocal localized function assignment | Aspose.Cells workbook culture pt-PT | Portuguese Excel formula evaluation .NET | semicolon separator in Aspose.Cells formulas | save workbook after FormulaLocal calculation

using System;
using System.Globalization;
using Aspose.Cells;

// The example creates a new workbook, sets its culture to Portuguese (Portugal), assigns the localized SUM function (SOMA) to cell A1 using FormulaLocal with semicolon separators, recalculates the workbook, and saves it as LocalizedFormula.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Set workbook culture to Portuguese (Portugal) so that localized formulas are recognized
            workbook.Settings.CultureInfo = new CultureInfo("pt-PT");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a formula using the localized function name.
            // In Portuguese Excel, the SUM function is "SOMA" and arguments are separated by ';'.
            Cell targetCell = sheet.Cells["A1"];
            targetCell.FormulaLocal = "=SOMA(10;20)";

            // Recalculate the workbook to evaluate the formula.
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule: save)
            string outputPath = "LocalizedFormula.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
