using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the workbook region to a locale (e.g., German) to work with localized formulas
        workbook.Settings.Region = CountryCode.Germany;

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Put sample values that the formula will use
        cells["B1"].PutValue(5);
        cells["C1"].PutValue(7);

        // Set a formula using the locale‑specific syntax (German SUMME)
        cells["A1"].FormulaLocal = "=SUMME(B1:C1)";

        // Enable automatic recalculation so that formulas are evaluated after changes
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Optional: ensure calculation occurs when the workbook is opened or saved
        workbook.Settings.FormulaSettings.CalculateOnOpen = true;
        workbook.Settings.FormulaSettings.CalculateOnSave = true;

        // Perform calculation immediately to obtain the result
        workbook.CalculateFormula();

        // Display the calculated value
        Console.WriteLine("A1 result: " + cells["A1"].Value);

        // Save the workbook
        workbook.Save("LocalizedFormula.xlsx");
    }
}