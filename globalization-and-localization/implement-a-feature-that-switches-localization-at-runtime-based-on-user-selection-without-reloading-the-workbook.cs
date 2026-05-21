using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace LocalizationDemo
{
    class LocalizationSwitcher
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue(10);
                ws.Cells["A2"].PutValue(20);
                ws.Cells["A3"].Formula = "=SUM(A1:A2)";

                // Calculate with the default (English) locale
                wb.CalculateFormula();
                Console.WriteLine($"Result with default locale: {ws.Cells["A3"].Value}");

                // Switch localization to French at runtime (no reload required)
                SwitchLocalization(wb, CountryCode.France);

                // Re‑calculate to reflect any locale‑dependent changes
                wb.CalculateFormula();
                Console.WriteLine($"Result after switching to French: {ws.Cells["A3"].Value}");

                // Map the standard function name "SUM" to its French equivalent "SOMME"
                SetLocalizedFunction(wb, "SUM", "SOMME");

                // Use the localized function name in a new formula
                ws.Cells["B1"].Formula = "=SOMME(A1:A2)";
                wb.CalculateFormula();
                Console.WriteLine($"Result using localized function name 'SOMME': {ws.Cells["B1"].Value}");

                // Save the workbook (optional)
                string outputPath = "LocalizationDemo.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Changes the workbook's UI language, region and culture info at runtime
        static void SwitchLocalization(Workbook workbook, CountryCode country)
        {
            // Set UI language and regional settings
            workbook.Settings.LanguageCode = country;
            workbook.Settings.Region = country;

            // Determine appropriate culture name for the given CountryCode
            string cultureName = GetCultureName(country);
            if (!string.IsNullOrEmpty(cultureName))
            {
                workbook.Settings.CultureInfo = new CultureInfo(cultureName);
            }

            // Ensure globalization settings are replaceable
            workbook.Settings.GlobalizationSettings = new SettableGlobalizationSettings();
        }

        // Returns a valid .NET culture name for a given CountryCode
        static string GetCultureName(CountryCode country)
        {
            switch (country)
            {
                case CountryCode.France:
                    return "fr-FR";
                case CountryCode.Germany:
                    return "de-DE";
                case CountryCode.Spain:
                    return "es-ES";
                case CountryCode.UnitedKingdom:
                    return "en-GB";
                // Add more mappings as needed
                default:
                    return CultureInfo.InvariantCulture.Name;
            }
        }

        // Adds a bidirectional mapping between a standard function name and a localized name
        static void SetLocalizedFunction(Workbook workbook, string standardName, string localName)
        {
            // Ensure the workbook uses SettableGlobalizationSettings
            var settings = workbook.Settings.GlobalizationSettings as SettableGlobalizationSettings;
            if (settings == null)
            {
                settings = new SettableGlobalizationSettings();
                workbook.Settings.GlobalizationSettings = settings;
            }

            // Map the standard function to the local name (bidirectional = true)
            settings.SetLocalFunctionName(standardName, localName, true);
        }
    }
}