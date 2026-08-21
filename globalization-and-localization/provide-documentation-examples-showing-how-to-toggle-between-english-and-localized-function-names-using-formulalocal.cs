// Title: Toggle English and German Excel function names with FormulaLocal in Aspose.Cells for .NET
// Description: Demonstrates how to set a workbook region to Germany, map the English function SUM to the German name SUMME using SettableGlobalizationSettings, and switch between standard Formula and localized FormulaLocal. The example writes sample data, assigns formulas in both languages, recalculates, and saves the workbook, showing bidirectional conversion of formulas.
// Keywords: Aspose.Cells | FormulaLocal | C# | .NET | Excel localization | German function names | SUMME | SettableGlobalizationSettings | globalization settings | toggle formulas | regional workbook
// Common Searches: Aspose.Cells FormulaLocal German example | map English SUM to SUMME Aspose.Cells | switch between English and localized formulas .NET | set workbook region Germany Aspose.Cells | bidirectional function name mapping Aspose
// Developer Intent: The developer needs to convert Excel formulas between English and a localized language (e.g., German) and retrieve both the standard and localized representations using Aspose.Cells.
// Use Cases: Create workbooks for German‑speaking users while preserving English formulas for internal processing. | Allow end‑users to enter formulas in their native language and automatically translate them to the engine’s standard English syntax. | Generate regional reports that display localized function names when opened in Excel, yet remain compatible with formula evaluation APIs.
// AI Prompts: Show how to map multiple Excel functions to French equivalents with SettableGlobalizationSettings in Aspose.Cells. | Provide a code snippet that updates the workbook region at runtime and automatically toggles formulas between English and the new locale. | Explain how to retrieve the original English formula after a user entered a localized formula via FormulaLocal.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to set a workbook region to Germany, map the English function SUM to the German name SUMME using SettableGlobalizationSettings, and switch between standard Formula and localized FormulaLocal. The example writes sample data, assigns formulas in both languages, recalculates, and saves the workbook, showing bidirectional conversion of formulas.
    public class ToggleFormulaLocalizationDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set the workbook region to German to demonstrate localization
                workbook.Settings.Region = CountryCode.Germany;

                // Create customizable globalization settings
                SettableGlobalizationSettings gSettings = new SettableGlobalizationSettings();

                // Map the standard English function name "SUM" to the German localized name "SUMME"
                // bidirectional = true enables automatic reverse mapping
                gSettings.SetLocalFunctionName("SUM", "SUMME", true);

                // Apply the globalization settings to the workbook
                workbook.Settings.GlobalizationSettings = gSettings;

                // Populate some sample data for the SUM calculation
                sheet.Cells["B1"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);
                sheet.Cells["B3"].PutValue(30);

                // Access cell A1 and set a formula using the standard (English) name
                Cell cell = sheet.Cells["A1"];
                cell.Formula = "=SUM(B1:B3)";

                // Display the formula in both standard and localized forms
                Console.WriteLine("After setting Formula (English):");
                Console.WriteLine("Standard Formula   : " + cell.Formula);
                Console.WriteLine("Localized Formula  : " + cell.FormulaLocal);

                // Now set the formula using the localized (German) name via FormulaLocal
                cell.FormulaLocal = "=SUMME(B1:B3)";

                // Display the formulas again to show the toggle effect
                Console.WriteLine("\nAfter setting FormulaLocal (German):");
                Console.WriteLine("Standard Formula   : " + cell.Formula);
                Console.WriteLine("Localized Formula  : " + cell.FormulaLocal);

                // Calculate the result to verify both formulas produce the same value
                workbook.CalculateFormula();
                Console.WriteLine("\nCalculated Value in A1: " + cell.Value);

                // Save the workbook (lifecycle rule: save)
                string outputPath = "ToggleFormulaLocalizationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during execution: " + ex.Message);
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unhandled exception: " + ex.Message);
            }
        }
    }
}
