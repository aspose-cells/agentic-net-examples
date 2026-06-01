using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Custom globalization settings that return locale‑specific boolean strings
    public class CustomBooleanGlobalizationSettings : GlobalizationSettings
    {
        private readonly CultureInfo _culture;

        public CustomBooleanGlobalizationSettings(CultureInfo culture)
        {
            _culture = culture;
        }

        // Override to provide localized true/false strings
        public override string GetBooleanValueString(bool bv)
        {
            // French example
            if (_culture.Name.StartsWith("fr", StringComparison.OrdinalIgnoreCase))
                return bv ? "Vrai" : "Faux";

            // German example
            if (_culture.Name.StartsWith("de", StringComparison.OrdinalIgnoreCase))
                return bv ? "Wahr" : "Falsch";

            // Default to English
            return bv ? "True" : "False";
        }
    }

    // Demonstration of applying the custom settings to a workbook
    public class BooleanLocalizationDemo
    {
        public static void Run()
        {
            try
            {
                // Select the desired locale (e.g., French - France)
                CultureInfo locale = new CultureInfo("fr-FR");

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Assign the custom globalization settings to the workbook
                workbook.Settings.GlobalizationSettings = new CustomBooleanGlobalizationSettings(locale);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Insert boolean values into cells
                sheet.Cells["A1"].PutValue(true);
                sheet.Cells["A2"].PutValue(false);

                // Save the workbook (the boolean strings will appear according to the locale)
                string outputPath = "BooleanLocalizationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point required for console application
    public class Program
    {
        public static void Main(string[] args)
        {
            BooleanLocalizationDemo.Run();
        }
    }
}