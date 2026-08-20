// Title: C# Tool to Compare English and Localized Excel Formulas with Aspose.Cells
// Description: Loads an Excel file, sets a target region (e.g., Germany), scans every formula cell and prints the standard English formula (Formula) together with its locale‑specific version (FormulaLocal). The sample also extracts the localized function name and converts it to the English equivalent via Aspose.Cells globalization settings, then optionally saves the workbook.
// Keywords: Aspose.Cells | C# formula localization | Formula vs FormulaLocal | Excel function translation | globalization settings | localized Excel formulas | QA Excel multilingual | German Excel functions
// Common Searches: compare Formula and FormulaLocal Aspose.Cells C# | map German Excel function name to English with Aspose | list localized formulas in a workbook using Aspose.Cells | verify Excel formula translation programmatically | extract function name from Excel formula C#
// Developer Intent: Provide a quick diagnostic utility that enumerates each cell’s English formula and its localized counterpart, and optionally resolves the localized function name back to the standard English name.
// Use Cases: Validate that a German‑localized workbook uses the correct English functions before release. | Generate a QA report showing cell address, English formula, localized formula, and mapped standard function. | Detect unsupported or mismatched localized functions across multilingual Excel templates.
// AI Prompts: Create C# code that logs differences between cell.Formula and cell.FormulaLocal for every formula cell in a workbook. | Demonstrate how to use Aspose.Cells.GlobalizationSettings.GetStandardFunctionName to translate a German function like 'SUMME' to its English name. | Write a method that exports cell address, English formula, localized formula, and mapped standard function to a CSV file.

using System;
using Aspose.Cells;

// Loads an Excel file, sets a target region (e.g., Germany), scans every formula cell and prints the standard English formula (Formula) together with its locale‑specific version (FormulaLocal). The sample also extracts the localized function name and converts it to the English equivalent via Aspose.Cells globalization settings, then optionally saves the workbook.
class FormulaLocalizationDiagnostic
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx"); // load rule

        // Set a locale to demonstrate localized formulas (e.g., German)
        workbook.Settings.Region = CountryCode.Germany;

        Worksheet worksheet = workbook.Worksheets[0];

        // Determine the used range of the worksheet
        int maxRow = worksheet.Cells.MaxDataRow;
        int maxCol = worksheet.Cells.MaxDataColumn;

        // Iterate through all cells that contain data
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = worksheet.Cells[row, col];

                // Process only cells that have a formula
                if (cell.IsFormula)
                {
                    // English (standard) formula
                    string englishFormula = cell.Formula;

                    // Localized formula according to the workbook's region
                    string localFormula = cell.FormulaLocal;

                    Console.WriteLine($"Cell {cell.Name}:");
                    Console.WriteLine($"  English Formula : {englishFormula}");
                    Console.WriteLine($"  Localized Formula: {localFormula}");

                    // Optional verification: map the localized function name back to the standard name
                    string localFunc = ExtractFunctionName(localFormula);
                    if (!string.IsNullOrEmpty(localFunc))
                    {
                        string standardFunc = workbook.Settings.GlobalizationSettings.GetStandardFunctionName(localFunc);
                        Console.WriteLine($"  Mapped Standard Function: {standardFunc}");
                    }
                }
            }
        }

        // Save the workbook after processing (optional)
        workbook.Save("output.xlsx"); // save rule
    }

    // Helper method to extract the function name from a formula string (e.g., "=SUMME(A1:A2)" -> "SUMME")
    static string ExtractFunctionName(string formula)
    {
        if (string.IsNullOrEmpty(formula))
            return null;

        // Remove leading '=' if present
        string trimmed = formula.TrimStart('=');

        // Function name ends at the first '(' character
        int idx = trimmed.IndexOf('(');
        if (idx > 0)
            return trimmed.Substring(0, idx);

        // If no '(' found, the whole string is considered the function name
        return trimmed;
    }
}
