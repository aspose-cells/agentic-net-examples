using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ToggleFormulaLocalizationDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set the workbook region to Germany to demonstrate German localization
                workbook.Settings.Region = CountryCode.Germany;

                // Create customizable globalization settings
                SettableGlobalizationSettings gSettings = new SettableGlobalizationSettings();

                // Map the standard English function name "SUM" to the German name "SUMME"
                // The bidirectional flag makes the mapping work both ways
                gSettings.SetLocalFunctionName("SUM", "SUMME", true);

                // Apply the custom globalization settings to the workbook
                workbook.Settings.GlobalizationSettings = gSettings;

                // ------------------------------------------------------------
                // 1. Use the standard (English) function name in a formula
                // ------------------------------------------------------------
                Cell cellEnglish = sheet.Cells["A1"];
                cellEnglish.Formula = "=SUM(B1:B3)";

                // Populate the source range
                sheet.Cells["B1"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);
                sheet.Cells["B3"].PutValue(30);

                Console.WriteLine("Standard Formula (English): " + cellEnglish.Formula);
                Console.WriteLine("Localized Formula (German): " + cellEnglish.FormulaLocal);

                // ------------------------------------------------------------
                // 2. Use the localized (German) function name via FormulaLocal
                // ------------------------------------------------------------
                Cell cellLocalized = sheet.Cells["A2"];
                cellLocalized.FormulaLocal = "=SUMME(C1:C3)";

                // Populate the source range for the second formula
                sheet.Cells["C1"].PutValue(5);
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);

                Console.WriteLine("Localized Formula (German) set via FormulaLocal: " + cellLocalized.FormulaLocal);
                Console.WriteLine("Standard Formula (English) derived from localized: " + cellLocalized.Formula);

                // ------------------------------------------------------------
                // Calculate all formulas to verify correctness
                // ------------------------------------------------------------
                workbook.CalculateFormula();

                Console.WriteLine("Result of A1 (SUM): " + cellEnglish.Value);
                Console.WriteLine("Result of A2 (SUMME): " + cellLocalized.Value);

                // Save the workbook to verify the formulas are stored correctly
                string outputPath = "ToggleFormulaLocalizationDemo.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ToggleFormulaLocalizationDemo.Run();
        }
    }
}