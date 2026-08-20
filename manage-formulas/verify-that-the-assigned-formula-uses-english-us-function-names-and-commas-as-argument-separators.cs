// Title: Check English (US) Formula Syntax and Comma Separators with Aspose.Cells for .NET
// Description: Shows how to set a workbook region (e.g., Germany), assign an English‑US formula (=SUM(1,2,3)) to a cell, verify that the Formula property uses English function names and commas, compare it with FormulaLocal, map German function names via GlobalizationSettings.GetStandardFunctionName, and save the workbook.
// Keywords: Aspose.Cells formula verification | English formula syntax | Formula vs FormulaLocal | GlobalizationSettings GetStandardFunctionName | workbook region Germany | .NET Excel localization | cross‑locale formulas
// Common Searches: Aspose.Cells verify English formula | Formula vs FormulaLocal in .NET | GetStandardFunctionName example C# | Set workbook region Germany Aspose.Cells | Check formula separators Aspose.Cells | Excel formula localization C#
// Developer Intent: Confirm that a formula assigned in code is stored in the standard English format with commas, regardless of the workbook's locale.
// Use Cases: Validate that automatically generated formulas remain in English syntax for cross‑regional Excel files. | Display both the standard Formula and the localized FormulaLocal values to users in different locales. | Convert localized function names (e.g., German "SUMME") to their English equivalents using GlobalizationSettings.GetStandardFunctionName.
// AI Prompts: Write C# code with Aspose.Cells that sets the workbook region to France and checks that a formula uses semicolons in the localized representation. | Create a method that receives any cell address and returns true if its Formula uses English function names and commas, using Aspose.Cells. | Explain how GlobalizationSettings.GetStandardFunctionName works and provide examples for German, French, and Spanish Excel function names.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to set a workbook region (e.g., Germany), assign an English‑US formula (=SUM(1,2,3)) to a cell, verify that the Formula property uses English function names and commas, compare it with FormulaLocal, map German function names via GlobalizationSettings.GetStandardFunctionName, and save the workbook.
    public class VerifyEnglishFormula
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Set workbook region to a non‑English locale (e.g., Germany) to demonstrate localization
                workbook.Settings.Region = CountryCode.Germany;

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Assign a formula using English function name and commas
                // This will be stored in the standard (en‑US) format regardless of the workbook locale
                cells["A1"].Formula = "=SUM(1,2,3)";

                // Retrieve the standard formula (English) and the localized formula
                string standardFormula = cells["A1"].Formula;          // Expected: "=SUM(1,2,3)"
                string localizedFormula = cells["A1"].FormulaLocal;    // May be "=SUMME(1;2;3)" in German locale

                // Verify that the standard formula uses English function name and commas
                bool usesEnglishFunction = standardFormula.Contains("SUM");
                bool usesCommas = standardFormula.Contains(",");

                Console.WriteLine("Standard Formula: " + standardFormula);
                Console.WriteLine("Localized Formula: " + localizedFormula);
                Console.WriteLine("Uses English function name: " + usesEnglishFunction);
                Console.WriteLine("Uses commas as separators: " + usesCommas);

                // Additional check: convert a German function name to its standard English name
                // (demonstrates GlobalizationSettings.GetStandardFunctionName)
                string germanFunction = "SUMME";
                string converted = workbook.Settings.GlobalizationSettings.GetStandardFunctionName(germanFunction);
                Console.WriteLine($"German function '{germanFunction}' maps to standard '{converted}'");

                // Save the workbook (lifecycle rule: save)
                string outputPath = "VerifyEnglishFormula.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public static class Program
    {
        public static void Main()
        {
            VerifyEnglishFormula.Run();
        }
    }
}
