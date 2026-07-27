using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Custom globalization settings that return locale‑specific boolean strings
    class CustomBooleanGlobalizationSettings : GlobalizationSettings
    {
        private readonly CultureInfo _culture;

        public CustomBooleanGlobalizationSettings(CultureInfo culture)
        {
            _culture = culture;
        }

        // Override to provide localized true/false representations
        public override string GetBooleanValueString(bool bv)
        {
            // French example
            if (_culture.Name.StartsWith("fr", StringComparison.OrdinalIgnoreCase))
                return bv ? "Vrai" : "Faux";

            // German example
            if (_culture.Name.StartsWith("de", StringComparison.OrdinalIgnoreCase))
                return bv ? "Wahr" : "Falsch";

            // Fallback to default English strings
            return bv ? "TRUE" : "FALSE";
        }
    }

    public class BooleanLocalizationDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Select the desired locale (e.g., French - France)
            CultureInfo locale = new CultureInfo("fr-FR");

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomBooleanGlobalizationSettings(locale);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert boolean values into cells
            sheet.Cells["A1"].PutValue(true);
            sheet.Cells["A2"].PutValue(false);

            // Save the workbook
            string outputPath = "BooleanLocalizationDemo.xlsx";
            workbook.Save(outputPath);

            // Demonstrate the overridden GetBooleanValueString method
            GlobalizationSettings gs = workbook.Settings.GlobalizationSettings;
            Console.WriteLine($"Localized true string: {gs.GetBooleanValueString(true)}");
            Console.WriteLine($"Localized false string: {gs.GetBooleanValueString(false)}");
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }

        // Entry point required by the runtime
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}