using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceLogging
{
    class Program
    {
        static void Main()
        {
            // Prepare a list of locales to be applied (localization steps)
            var locales = new List<string> { "en", "de", "fr", "es", "zh" };

            // Create a workbook (creation rule)
            Workbook workbook = new Workbook();

            // Add some sample data and a formula that will be used after localization
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["B1"].PutValue(5);
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(15);
            sheet.Cells["A1"].Formula = "=SUM(B1:B3)";

            // Container for performance results
            var performanceLog = new List<string>();

            // Iterate through each localization step
            foreach (string locale in locales)
            {
                // Start timing this localization step
                Stopwatch sw = Stopwatch.StartNew();

                // Apply localization settings
                SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
                // Example: map the built‑in function name "SUM" to a locale‑specific name.
                // In a real scenario the local name would be the actual translation.
                string localSumName = GetLocalSumName(locale);
                globalization.SetLocalBuiltInName("SUM", localSumName, true);
                workbook.Settings.GlobalizationSettings = globalization;

                // Force recalculation to ensure the formula uses the new settings
                workbook.CalculateFormula();

                // Stop timing
                sw.Stop();

                // Record the elapsed time for this locale
                performanceLog.Add($"Locale: {locale}, TimeMs: {sw.ElapsedMilliseconds}");
            }

            // Output performance metrics
            Console.WriteLine("Localization Performance Metrics:");
            foreach (string entry in performanceLog)
            {
                Console.WriteLine(entry);
            }

            // Save the workbook (save rule)
            workbook.Save("LocalizedResult.xlsx");
        }

        // Helper method to return a mock local name for the SUM function based on locale.
        // In practice you would use the actual translated function name.
        static string GetLocalSumName(string locale)
        {
            return locale switch
            {
                "de" => "SUMME",      // German
                "fr" => "SOMME",      // French
                "es" => "SUMA",       // Spanish
                "zh" => "求和",        // Chinese
                _ => "SUM",           // Default (English)
            };
        }
    }
}