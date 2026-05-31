using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom globalization settings that fall back to English (standard) function names
    // when a locale‑specific mapping is not defined.
    public class FallbackGlobalizationSettings : GlobalizationSettings
    {
        // Mapping from standard (English) function names to localized names.
        private readonly Dictionary<string, string> _standardToLocal = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "SUM", "LOCALSUM" } // Example mapping; add more as needed.
        };

        // Reverse lookup dictionary.
        private readonly Dictionary<string, string> _localToStandard;

        public FallbackGlobalizationSettings()
        {
            _localToStandard = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in _standardToLocal)
                _localToStandard[kvp.Value] = kvp.Key;
        }

        // Returns the localized function name if defined; otherwise falls back to the standard name.
        public override string GetLocalFunctionName(string standardName)
        {
            if (standardName == null) throw new ArgumentNullException(nameof(standardName));

            return _standardToLocal.TryGetValue(standardName, out string localName) ? localName : standardName;
        }

        // Returns the standard function name for a given localized name.
        // If the localized name is not mapped, assume it is already the standard English name.
        public override string GetStandardFunctionName(string localName)
        {
            if (localName == null) throw new ArgumentNullException(nameof(localName));

            return _localToStandard.TryGetValue(localName, out string standardName) ? standardName : localName;
        }
    }

    public class FallbackLocalizationDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Apply the custom globalization settings.
                workbook.Settings.GlobalizationSettings = new FallbackGlobalizationSettings();

                // Populate sample data.
                sheet.Cells["B1"].PutValue(5);
                sheet.Cells["B2"].PutValue(15);

                // Use the localized function name in a formula.
                // NOTE: If the custom globalization mapping does not work in a particular environment,
                // fall back to the standard function name to avoid runtime errors.
                Cell localizedCell = sheet.Cells["B3"];
                localizedCell.Formula = "=LOCALSUM(B1:B2)";

                // Use the standard English function name in another formula.
                Cell standardCell = sheet.Cells["B4"];
                standardCell.Formula = "=SUM(B1:B2)";

                // Calculate all formulas.
                workbook.CalculateFormula();

                // Output results.
                Console.WriteLine($"Result using localized name (LOCALSUM): {localizedCell.DoubleValue}");
                Console.WriteLine($"Result using standard name (SUM): {standardCell.DoubleValue}");

                // Save the workbook (optional).
                string outputPath = "FallbackLocalizationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point.
    class Program
    {
        static void Main()
        {
            FallbackLocalizationDemo.Run();
        }
    }
}