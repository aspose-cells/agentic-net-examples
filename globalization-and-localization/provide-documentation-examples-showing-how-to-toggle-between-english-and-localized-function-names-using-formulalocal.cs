// Title: Toggle English ↔ German formulas using FormulaLocal in Aspose.Cells for .NET
// Description: A C# demo that creates a workbook, sets its region to Germany, adds a bidirectional mapping from the English function SUM to the German SUMME via SettableGlobalizationSettings, and illustrates using both the standard Formula property and the localized FormulaLocal property. After calculation the results match, and the workbook is saved with both formula representations.
// Keywords: Aspose.Cells | .NET | FormulaLocal | formula localization | English German formulas | SettableGlobalizationSettings | custom globalization | SUM to SUMME | region Germany | toggle formula language | localized functions
// Common Searches: Aspose.Cells FormulaLocal German example | map English SUM to German SUMME Aspose.Cells | toggle between English and localized formulas .NET | set workbook region Germany Aspose.Cells | custom globalization settings C# Aspose.Cells
// Developer Intent: Demonstrate switching between English and German formulas in an Aspose.Cells workbook using FormulaLocal and a custom globalization mapping.
// Use Cases: Generate spreadsheets for German users by mapping English functions to their German equivalents and entering formulas with FormulaLocal. | Read workbooks containing localized formulas while accessing the standard English syntax via the Formula property. | Produce multi‑region reports where users input formulas in their native language but calculations are performed uniformly.
// AI Prompts: Create C# code that maps the French function 'SOMME' to English 'SUM' using SettableGlobalizationSettings and shows FormulaLocal usage. | Provide an Aspose.Cells example that toggles between English and Spanish formulas with FormulaLocal and verifies the results. | Explain how to preserve both English and localized formulas when saving a workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A C# demo that creates a workbook, sets its region to Germany, adds a bidirectional mapping from the English function SUM to the German SUMME via SettableGlobalizationSettings, and illustrates using both the standard Formula property and the localized FormulaLocal property. After calculation the results match, and the workbook is saved with both formula representations.
    public class ToggleFormulaLocalizationDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set the workbook region to German to demonstrate localization
                workbook.Settings.Region = CountryCode.Germany;

                // Create customizable globalization settings
                SettableGlobalizationSettings settings = new SettableGlobalizationSettings();

                // Map the standard English function name "SUM" to the German localized name "SUMME"
                // The bidirectional flag makes the mapping work both ways
                settings.SetLocalFunctionName("SUM", "SUMME", true);

                // Apply the custom globalization settings to the workbook
                workbook.Settings.GlobalizationSettings = settings;

                // ------------------------------------------------------------
                // 1. Use the standard (English) formula syntax
                // ------------------------------------------------------------
                // Fill some sample data
                sheet.Cells["B1"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);
                sheet.Cells["B3"].PutValue(30);

                // Set formula using the standard English name
                Cell cellEnglish = sheet.Cells["A1"];
                cellEnglish.Formula = "=SUM(B1:B3)";

                // Display both the standard and the localized representation
                Console.WriteLine("After setting Formula (English):");
                Console.WriteLine("Standard Formula   : " + cellEnglish.Formula);
                Console.WriteLine("Localized Formula  : " + cellEnglish.FormulaLocal);

                // ------------------------------------------------------------
                // 2. Use the localized (German) formula syntax via FormulaLocal
                // ------------------------------------------------------------
                // Change the formula using the localized name
                Cell cellLocalized = sheet.Cells["A2"];
                cellLocalized.FormulaLocal = "=SUMME(B1:B3)";

                // Display both representations again
                Console.WriteLine("\nAfter setting FormulaLocal (German):");
                Console.WriteLine("Standard Formula   : " + cellLocalized.Formula);
                Console.WriteLine("Localized Formula  : " + cellLocalized.FormulaLocal);

                // ------------------------------------------------------------
                // 3. Calculate formulas to verify they work identically
                // ------------------------------------------------------------
                workbook.CalculateFormula();

                Console.WriteLine("\nCalculated Results:");
                Console.WriteLine("A1 (English)  = " + cellEnglish.Value);
                Console.WriteLine("A2 (German)   = " + cellLocalized.Value);

                // Save the workbook to verify the formulas are stored correctly
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ToggleFormulaLocalizationDemo.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ToggleFormulaLocalizationDemo.Run();
        }
    }
}
