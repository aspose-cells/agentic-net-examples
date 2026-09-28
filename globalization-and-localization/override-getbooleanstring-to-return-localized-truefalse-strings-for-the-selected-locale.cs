// Title: Implement a custom Boolean string localizer and write locale‑specific true/false values to an Excel worksheet using Aspose.Cells in C#
// AI Prompts: Create a C# class named MyStringLocalizer with a GetBooleanString(bool) method that returns the appropriate true/false text for a given locale, then use this class to fill cells in an Aspose.Cells workbook with the localized strings. | Add additional culture cases to the GetBooleanString switch (e.g., it-IT, ja-JP) and modify the example to write those localized boolean strings into the worksheet before saving the file.
// Common Searches: how to display true/false in French using Aspose.Cells C# | C# Aspose.Cells write localized boolean text to cells | custom GetBooleanString implementation for multiple cultures in Aspose.Cells | example of culture‑specific boolean string conversion before saving Excel with Aspose.Cells | Aspose.Cells workbook localization of boolean values C#
// Tags: boolean string localization Aspose.Cells | custom culture formatter C# | write localized boolean values to Excel | extend GetBooleanString for multiple locales | Aspose.Cells workbook localization example

using System;
using Aspose.Cells;

namespace LocalizationExample
{
    // Simple string localizer that returns localized true/false strings based on the selected locale.
    // The example defines a MyStringLocalizer class that returns true/false strings based on a supplied locale, writes raw Boolean values to column A of a new workbook, writes the corresponding localized strings to column B, and saves the file as LocalizedBooleans.xlsx using Aspose.Cells.
    public class MyStringLocalizer
    {
        private readonly string _locale;

        // Constructor accepts a locale identifier (e.g., "en-US", "fr-FR").
        public MyStringLocalizer(string locale)
        {
            _locale = locale;
        }

        // Returns a localized representation of a Boolean value.
        public string GetBooleanString(bool value)
        {
            // Simple locale handling – extend this switch with more languages as needed.
            switch (_locale)
            {
                case "en-US":
                    return value ? "True" : "False";
                case "fr-FR":
                    return value ? "Vrai" : "Faux";
                case "de-DE":
                    return value ? "Wahr" : "Falsch";
                case "es-ES":
                    return value ? "Verdadero" : "Falso";
                case "zh-CN":
                    return value ? "真" : "假";
                default:
                    // Fallback to English if locale is not recognized.
                    return value ? "True" : "False";
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Initialize the custom string localizer (French locale for this example).
                MyStringLocalizer localizer = new MyStringLocalizer("fr-FR");

                // Write Boolean values to cells.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue(true);
                sheet.Cells["A2"].PutValue(false);

                // Retrieve the localized string representation of the Boolean values.
                string localizedTrue = localizer.GetBooleanString(true);   // "Vrai"
                string localizedFalse = localizer.GetBooleanString(false); // "Faux"

                // Write the localized strings back to the sheet for demonstration.
                sheet.Cells["B1"].PutValue(localizedTrue);
                sheet.Cells["B2"].PutValue(localizedFalse);

                // Save the workbook.
                string outputPath = "LocalizedBooleans.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
