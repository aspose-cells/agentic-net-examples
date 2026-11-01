// Title: How to override GetBooleanString in Aspose.Cells to return locale‑specific true/false strings in C#
// Common Searches: C# Aspose.Cells example for displaying true/false in French, German, or Spanish | How to customize boolean string conversion in an Aspose.Cells workbook using CultureInfo | Write localized boolean values to Excel cells with Aspose.Cells in C# | Create a custom BooleanLocalizer class for Aspose.Cells localization | Save an Excel file with localized true/false strings using Aspose.Cells
// Tags: custom boolean localization Aspose.Cells | GetBooleanString override C# | CultureInfo based boolean mapping Aspose.Cells | write localized boolean strings to Excel with Aspose.Cells | Aspose.Cells workbook localization example

using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsLocalizationExample
{
    // Helper class to provide locale‑specific boolean string representations
    // The sample defines a CustomBooleanLocalizer class that maps a boolean value to French, German, Spanish, or default English strings based on a supplied CultureInfo. It writes raw boolean values and their localized string equivalents to cells in a new Aspose.Cells workbook and saves the file, demonstrating how to override GetBooleanString for locale‑specific true/false output.
    public class CustomBooleanLocalizer
    {
        private readonly CultureInfo _culture;

        public CustomBooleanLocalizer(CultureInfo culture)
        {
            _culture = culture;
        }

        // Returns the localized string for a boolean value
        public string GetBooleanString(bool value)
        {
            // Example mappings for a few locales; extend as needed
            if (_culture.Name.StartsWith("fr", StringComparison.OrdinalIgnoreCase))
                return value ? "Vrai" : "Faux";

            if (_culture.Name.StartsWith("de", StringComparison.OrdinalIgnoreCase))
                return value ? "Wahr" : "Falsch";

            if (_culture.Name.StartsWith("es", StringComparison.OrdinalIgnoreCase))
                return value ? "Verdadero" : "Falso";

            // Default to English
            return value ? "True" : "False";
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Use the default globalization settings (no custom override needed)
                // workbook.Settings.GlobalizationSettings = workbook.Settings.GlobalizationSettings;

                // Write boolean values to cells
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue(true);
                sheet.Cells["A2"].PutValue(false);

                // Localize boolean strings using the helper class
                var localizer = new CustomBooleanLocalizer(new CultureInfo("fr-FR"));
                string trueString = localizer.GetBooleanString(true);
                string falseString = localizer.GetBooleanString(false);

                // Write the localized strings to adjacent cells for demonstration
                sheet.Cells["B1"].PutValue(trueString);
                sheet.Cells["B2"].PutValue(falseString);

                // Save the workbook (adjust path as needed)
                string outputPath = "LocalizedBooleans.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log or display the error details
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
