using System;
using Aspose.Cells;

class NonEnglishFormulaGuidelines
{
    static void Main()
    {
        // Guideline 1: Disable automatic formula parsing when loading the workbook.
        // This prevents parsing errors caused by formulas written in a different language.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = false; // skip parsing on open

        // Load the workbook (replace "input.xlsx" with the actual file path).
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Guideline 2: Map localized function names to their standard English equivalents
        // and enable locale‑dependent parsing so that formulas written in the local language are recognized.
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
        // Example for Italian: SUM -> SOMMA, AVERAGE -> MEDIA
        globalization.SetLocalFunctionName("SUM", "SOMMA", true);
        globalization.SetLocalFunctionName("AVERAGE", "MEDIA", true);
        workbook.Settings.GlobalizationSettings = globalization;

        // Parse all formulas that were loaded without parsing.
        workbook.ParseFormulas(false);

        // Optionally set the workbook region to match the user's locale (e.g., Italy).
        workbook.Settings.Region = CountryCode.Italy;

        // Calculate all formulas after parsing.
        workbook.CalculateFormula();

        // Save the processed workbook.
        workbook.Save("output.xlsx");
    }
}