// Title: Aspose.Cells .NET – Enable Automatic Recalculation for Localized Formulas (German Region)
// Description: A .NET sample that sets Workbook.Settings.Region to Germany, writes formulas with both the English `Formula` property and the German `FormulaLocal` property, switches the calculation mode to `CalcModeType.Automatic`, runs `CalculateFormula`, and saves the workbook. The code ensures localized formulas are evaluated instantly without extra manual steps.
// Keywords: Aspose.Cells | .NET | Workbook.Settings.Region | FormulaLocal | German Excel formulas | automatic calculation mode | CalcModeType.Automatic | localized formulas | globalization | Excel localization | auto recalc
// Common Searches: Aspose.Cells enable automatic calculation for FormulaLocal | set workbook region to Germany Aspose.Cells | how to recalculate localized formulas in .NET | CalcModeType Automatic Aspose.Cells example | FormulaLocal German function names Aspose.Cells
// Developer Intent: Configure a workbook so that localized formulas are calculated automatically, eliminating the need for explicit recalculation calls.
// Use Cases: Create a report that mixes English and German formulas and rely on automatic recalculation to keep totals correct. | Switch the workbook region to France, use French function names via FormulaLocal, and generate a ready‑to‑publish file without manual CalculateFormula. | Build a multi‑locale spreadsheet template where users can enter formulas in their native language and see instant results.
// AI Prompts: Show me how to set Workbook.Settings.Region to Japan and enable automatic recalculation for Japanese FormulaLocal expressions in Aspose.Cells .NET. | Provide code that changes the calculation mode to Manual, updates several localized formulas, and then triggers a single CalculateFormula call. | Explain how to verify that automatic recalculation works for a workbook containing both English and localized formulas after saving.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A .NET sample that sets Workbook.Settings.Region to Germany, writes formulas with both the English `Formula` property and the German `FormulaLocal` property, switches the calculation mode to `CalcModeType.Automatic`, runs `CalculateFormula`, and saves the workbook. The code ensures localized formulas are evaluated instantly without extra manual steps.
    public class EnableAutomaticRecalculationAfterLocalizedFormulas
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Set the workbook region to a locale (e.g., German) so that
                // localized formulas can be used via the FormulaLocal property.
                workbook.Settings.Region = CountryCode.Germany;

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Set a formula in the standard (English) format for reference
                cells["A1"].Formula = "=SUM(B1:C1)";

                // Set a localized formula (German) using the FormulaLocal property.
                // In German the SUM function is "SUMME".
                cells["A2"].FormulaLocal = "=SUMME(B2:C2)";

                // Populate the referenced cells with values
                cells["B1"].PutValue(10);
                cells["C1"].PutValue(20);
                cells["B2"].PutValue(5);
                cells["C2"].PutValue(15);

                // Enable automatic recalculation for the workbook.
                // This ensures that after setting localized formulas the
                // calculation engine runs automatically when CalculateFormula is called.
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

                // Perform calculation so that the results are up‑to‑date.
                workbook.CalculateFormula();

                // Output the calculated results to verify correctness.
                Console.WriteLine("Result of A1 (English formula): " + cells["A1"].IntValue); // Expected 30
                Console.WriteLine("Result of A2 (German formula): " + cells["A2"].IntValue); // Expected 20

                // Save the workbook (using the standard save rule)
                workbook.Save("LocalizedFormulas_AutoRecalc.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            EnableAutomaticRecalculationAfterLocalizedFormulas.Run();
        }
    }
}
