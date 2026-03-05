using System;
using Aspose.Cells;

namespace FormulaLocalLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (XLSX format)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Scenario 1: Default locale (en-US) – set formula in standard format
            // ------------------------------------------------------------
            Cell cellA1 = cells["A1"];
            cellA1.Formula = "=SUM(B1:C1)"; // English formula
            Console.WriteLine("Scenario 1 - Standard Formula: " + cellA1.Formula);
            Console.WriteLine("Scenario 1 - Localized Formula (default locale): " + cellA1.FormulaLocal);

            // ------------------------------------------------------------
            // Scenario 2: Change workbook region to German and observe localization
            // ------------------------------------------------------------
            workbook.Settings.Region = CountryCode.Germany; // Set locale to German
            Cell cellA2 = cells["A2"];
            cellA2.Formula = "=SUM(B2:C2)"; // Still using English function name
            Console.WriteLine("\nScenario 2 - After setting Region to Germany");
            Console.WriteLine("Standard Formula: " + cellA2.Formula);
            Console.WriteLine("Localized Formula: " + cellA2.FormulaLocal); // Should show German function name (SUMME)

            // ------------------------------------------------------------
            // Scenario 3: Set formula using the localized (German) syntax via FormulaLocal
            // ------------------------------------------------------------
            Cell cellA3 = cells["A3"];
            cellA3.FormulaLocal = "=SUMME(B3:C3)"; // German function name
            Console.WriteLine("\nScenario 3 - Set FormulaLocal with German syntax");
            Console.WriteLine("Standard Formula after setting FormulaLocal: " + cellA3.Formula);
            Console.WriteLine("Localized Formula: " + cellA3.FormulaLocal);

            // ------------------------------------------------------------
            // Scenario 4: Custom globalization – map a custom local function name
            // ------------------------------------------------------------
            SettableGlobalizationSettings customSettings = new SettableGlobalizationSettings();
            // Map standard "AVERAGE" to a fictional local name "MITTELWERT"
            customSettings.SetLocalFunctionName("AVERAGE", "MITTELWERT", true);
            workbook.Settings.GlobalizationSettings = customSettings;

            Cell cellA4 = cells["A4"];
            // Use the custom local function name in the formula
            cellA4.Formula = "=MITTELWERT(B4:C4)";
            Console.WriteLine("\nScenario 4 - Custom globalization with local function 'MITTELWERT'");
            Console.WriteLine("Standard Formula stored: " + cellA4.Formula);
            Console.WriteLine("Localized Formula: " + cellA4.FormulaLocal);

            // ------------------------------------------------------------
            // Scenario 5: Using FormulaParseOptions to input a locale‑dependent formula directly
            // ------------------------------------------------------------
            FormulaParseOptions options = new FormulaParseOptions
            {
                LocaleDependent = true, // Indicates the formula string is locale formatted
                R1C1Style = false
            };
            Cell cellA5 = cells["A5"];
            // French date format example; locale dependent flag tells Aspose to treat it as such
            cellA5.SetFormula("=TEXTE(AUJOURDHUI();\"[$-fr-FR]dddd, dd mmmm yyyy\")", options);
            Console.WriteLine("\nScenario 5 - Formula set with LocaleDependent = true");
            Console.WriteLine("Standard Formula: " + cellA5.Formula);
            Console.WriteLine("Localized Formula: " + cellA5.FormulaLocal);

            // ------------------------------------------------------------
            // Scenario 6: Retrieve formulas in both standard and localized forms
            // ------------------------------------------------------------
            Console.WriteLine("\nRetrieving formulas in both representations:");
            for (int row = 1; row <= 5; row++)
            {
                Cell c = cells[$"A{row}"];
                Console.WriteLine($"Row {row}: Standard = {c.Formula}, Localized = {c.FormulaLocal}");
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
        }
    }
}