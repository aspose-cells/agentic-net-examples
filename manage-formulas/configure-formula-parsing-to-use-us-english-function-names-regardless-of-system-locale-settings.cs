// Title: Force Aspose.Cells to parse formulas with US English function names (C#)
// Description: Demonstrates how to set a workbook's region to United States, apply default globalization settings, and use FormulaParseOptions with LocaleDependent = false so that English function names (e.g., SUM) are parsed and calculated correctly on any system locale, then saves the workbook.
// Keywords: Aspose.Cells | C# formula parsing | US English function names | FormulaParseOptions LocaleDependent | disable locale‑dependent formulas | set workbook region US | SettableGlobalizationSettings | locale‑independent calculation | force English formulas
// Common Searches: Aspose.Cells use English function names regardless of locale | C# set workbook region to US for formula parsing | disable locale dependent formula parsing Aspose.Cells | FormulaParseOptions LocaleDependent false example | force English formulas in non‑English Windows
// Developer Intent: Ensure that all formulas are interpreted with US English function names, independent of the operating system or workbook locale.
// Use Cases: Run a SUM formula written in English on a workbook created on a German‑language machine. | Generate reports that must retain consistent English formula syntax across global deployments. | Automate spreadsheet calculations in a multilingual environment without locale‑specific adjustments.
// AI Prompts: Show how to configure Aspose.Cells to always use US English function names when parsing formulas in C#. | Provide a C# snippet that sets FormulaParseOptions.LocaleDependent to false and evaluates a SUM formula. | Explain why setting workbook.Settings.Region to US and using SettableGlobalizationSettings makes formula parsing locale‑independent.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to force formula parsing to use US English function names
    // regardless of the system or workbook locale.
    // Demonstrates how to set a workbook's region to United States, apply default globalization settings, and use FormulaParseOptions with LocaleDependent = false so that English function names (e.g., SUM) are parsed and calculated correctly on any system locale, then saves the workbook.
    public class ForceEnglishFormulaParsing
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // 1. Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // 2. Set the workbook region explicitly to United States.
                // Use Enum.Parse to avoid compile‑time dependency on a specific enum member name.
                workbook.Settings.Region = (CountryCode)Enum.Parse(typeof(CountryCode), "US");

                // 3. Assign a SettableGlobalizationSettings instance.
                // No custom local function names are added, so only standard (English) names are recognized.
                workbook.Settings.GlobalizationSettings = new SettableGlobalizationSettings();

                // 4. Prepare some data for the formula.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);

                // 5. Create FormulaParseOptions with LocaleDependent = false (default).
                // This tells Aspose.Cells that the formula string is NOT locale‑formatted.
                FormulaParseOptions options = new FormulaParseOptions
                {
                    LocaleDependent = false // enforce English function names
                };

                // 6. Set a formula using the standard English function name "SUM".
                // The formula will be parsed correctly even if the system locale were, for example, German.
                Cell targetCell = sheet.Cells["B1"];
                targetCell.SetFormula("=SUM(A1:A3)", options);

                // 7. Calculate the workbook to evaluate the formula.
                workbook.CalculateFormula();

                // 8. Output the result to verify correct parsing.
                Console.WriteLine($"Result of English SUM formula: {targetCell.Value}");

                // 9. Save the workbook (lifecycle rule: save)
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ForceEnglishFormulaParsing.xlsx");
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}
