using System;
using Aspose.Cells;

class NonEnglishFormulaTips
{
    static void Main()
    {
        // Tip 1: Disable automatic formula parsing when loading an XLSX file.
        // This prevents parsing errors for formulas written in a non‑English locale.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = false; // use LoadOptions.ParsingFormulaOnOpen property
        Workbook workbook = new Workbook("input.xlsx", loadOptions); // load with options

        // Tip 2: Use locale‑aware handling when creating or modifying formulas.
        // Example: German locale where the SUM function is "SUMME".

        // Set the workbook region to German so that locale‑dependent features work correctly.
        workbook.Settings.Region = CountryCode.Germany;

        // Approach A: Use the FormulaLocal property to set a formula in the local language.
        Cell cellA1 = workbook.Worksheets[0].Cells["A1"];
        cellA1.FormulaLocal = "=SUMME(B1:C1)"; // use Cell.FormulaLocal property

        // Approach B: Use SetFormula with FormulaParseOptions.LocaleDependent = true.
        FormulaParseOptions parseOptions = new FormulaParseOptions
        {
            LocaleDependent = true // enable locale‑dependent parsing
        };
        Cell cellA2 = workbook.Worksheets[0].Cells["A2"];
        cellA2.SetFormula("=SUMME(B2:C2)", parseOptions, null); // use SetFormula(string, FormulaParseOptions, object)

        // Re‑calculate formulas after setting them.
        workbook.CalculateFormula(); // calculate all formulas

        // Save the modified workbook.
        workbook.Save("output.xlsx");
    }
}