using System;
using Aspose.Cells;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    // Custom globalization settings that fall back to English (standard) names
    // when a locale‑specific mapping is not defined.
    public class FallbackGlobalizationSettings : SettableGlobalizationSettings
    {
        // Return the localized name if it exists; otherwise return the standard name.
        public override string GetLocalFunctionName(string standardName)
        {
            string local = base.GetLocalFunctionName(standardName);
            // If the base implementation could not find a mapping it returns the same name.
            if (string.Equals(local, standardName, StringComparison.OrdinalIgnoreCase))
            {
                // Fallback to the English (standard) name.
                return standardName;
            }
            return local;
        }

        // Return the standard name if it exists; otherwise assume the supplied name is already English.
        public override string GetStandardFunctionName(string localName)
        {
            string standard = base.GetStandardFunctionName(localName);
            if (string.Equals(standard, localName, StringComparison.OrdinalIgnoreCase))
            {
                // No locale mapping – treat the provided name as the English name.
                return localName;
            }
            return standard;
        }
    }

    public class FallbackFunctionNameDemo
    {
        public static void Run()
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare sample data.
            sheet.Cells["B1"].PutValue(5);
            sheet.Cells["B2"].PutValue(15);
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["B4"].PutValue(35);

            // Create custom globalization settings.
            FallbackGlobalizationSettings settings = new FallbackGlobalizationSettings();

            // Map only the SUM function to a locale‑specific name.
            // All other functions will rely on the fallback logic.
            settings.SetLocalFunctionName("SUM", "LOCALSUM", true);

            // Apply the settings to the workbook.
            workbook.Settings.GlobalizationSettings = settings;

            // Use the localized name – this works because we defined the mapping.
            Cell cellLocalizedSum = sheet.Cells["C1"];
            cellLocalizedSum.Formula = "=LOCALSUM(B1:B4)";

            // Use a standard English function name – works without any mapping.
            Cell cellStandardAvg = sheet.Cells["C2"];
            cellStandardAvg.Formula = "=AVERAGE(B1:B4)";

            // Attempt to use a localized name that has no mapping.
            // The fallback will treat "LOCALAVERAGE" as the English name "AVERAGE".
            Cell cellFallbackAvg = sheet.Cells["C3"];
            // Manually obtain the standard name using the fallback logic.
            string standardForLocalAvg = settings.GetStandardFunctionName("LOCALAVERAGE");
            // Build the formula with the resolved standard name.
            cellFallbackAvg.Formula = $"={standardForLocalAvg}(B1:B4)";

            // Calculate all formulas.
            workbook.CalculateFormula();

            // Output results.
            Console.WriteLine($"Result of LOCALSUM (localized): {cellLocalizedSum.Value}");
            Console.WriteLine($"Result of AVERAGE (standard): {cellStandardAvg.Value}");
            Console.WriteLine($"Result of LOCALAVERAGE (fallback to standard): {cellFallbackAvg.Value}");

            // Save the workbook.
            workbook.Save("FallbackFunctionNameDemo.xlsx");
        }
    }

    // Entry point for demonstration.
    class Program
    {
        static void Main()
        {
            FallbackFunctionNameDemo.Run();
        }
    }
}