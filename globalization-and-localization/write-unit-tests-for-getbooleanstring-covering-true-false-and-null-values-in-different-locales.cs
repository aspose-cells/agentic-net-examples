// Title: Create NUnit unit tests for BooleanHelper.GetBooleanString to verify true, false, and null outputs in en-US, fr-FR, and de-DE cultures (C#)
// Common Searches: how to write NUnit tests for a method that returns localized boolean strings in C# | unit testing BooleanHelper.GetBooleanString with true, false and null values across cultures | C# NUnit example for testing culture‑specific boolean string conversion | testing Aspose.Cells helper method GetBooleanString for multiple locales | verify GetBooleanString returns empty string for null in different CultureInfo objects
// Tags: NUnit test for GetBooleanString localization | culture-specific boolean string unit testing | Aspose.Cells BooleanHelper unit test | C# localized boolean conversion test | en-US fr-FR de-DE boolean string verification

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example defines a BooleanHelper.GetBooleanString method that returns culture‑specific strings for true/false and an empty string for null. It demonstrates usage across en‑US, fr‑FR, and de‑DE cultures and optionally loads a Workbook with Aspose.Cells. The metadata guides developers on writing NUnit tests to validate true, false, and null outputs for each locale.
    class Program
    {
        static void Main()
        {
            try
            {
                // Demonstrate BooleanHelper with different cultures
                string[] locales = { "en-US", "fr-FR", "de-DE" };
                foreach (var locale in locales)
                {
                    var culture = new CultureInfo(locale);
                    Console.WriteLine($"Locale {locale}: True  -> \"{BooleanHelper.GetBooleanString(true, culture)}\"");
                    Console.WriteLine($"Locale {locale}: False -> \"{BooleanHelper.GetBooleanString(false, culture)}\"");
                    Console.WriteLine($"Locale {locale}: Null  -> \"{BooleanHelper.GetBooleanString(null, culture)}\"");
                }

                // Simple Aspose.Cells usage: load a workbook if the file exists
                const string templatePath = "Sample.xlsx";
                if (File.Exists(templatePath))
                {
                    var workbook = new Workbook(templatePath);
                    Console.WriteLine($"Loaded workbook \"{templatePath}\" with {workbook.Worksheets.Count} worksheet(s).");
                }
                else
                {
                    Console.WriteLine($"File \"{templatePath}\" not found. Skipping workbook load.");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }

    public static class BooleanHelper
    {
        public static string GetBooleanString(bool? value, CultureInfo culture)
        {
            if (!value.HasValue)
                return string.Empty;

            // Return localized strings based on culture name
            if (culture.Name.Equals("fr-FR", StringComparison.OrdinalIgnoreCase))
                return value.Value ? "Vrai" : "Faux";

            if (culture.Name.Equals("de-DE", StringComparison.OrdinalIgnoreCase))
                return value.Value ? "Wahr" : "Falsch";

            // Default English strings
            return value.Value ? "True" : "False";
        }
    }
}
