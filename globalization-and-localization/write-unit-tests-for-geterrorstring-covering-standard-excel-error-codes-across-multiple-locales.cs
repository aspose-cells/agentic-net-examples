// Title: Create C# unit tests for Aspose.Cells CellsHelper.GetErrorString to verify Excel error strings across en-US, fr-FR, de-DE, ja-JP, and zh-CN locales
// AI Prompts: Generate an MSTest or NUnit parameterized test that iterates over each Aspose.Cells.ErrorValueType value, sets Thread.CurrentThread.CurrentCulture to a target locale, invokes CellsHelper.GetErrorString via reflection, and asserts the returned string matches the expected localized error text. | Write a reusable test helper method that accepts a culture name and an error enum name, switches the thread culture, calls CellsHelper.GetErrorString, and returns the result for data‑driven verification of all standard Excel error codes.
// Common Searches: how to unit test Aspose.Cells GetErrorString for different cultures in C# | verify localized Excel error messages with Aspose.Cells .NET | C# test for Aspose.Cells error enum translations across en-US and fr-FR | parameterized unit test for Excel error codes using Aspose.Cells CellsHelper | reflection based testing of Aspose.Cells GetErrorString method
// Tags: Aspose.Cells GetErrorString localization unit test | C# reflection CellsHelper error string verification | Excel error code culture testing .NET | parameterized Aspose.Cells error enum tests | multilingual Excel error string validation

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The sample uses reflection to locate the Aspose.Cells.ErrorValueType enum and the CellsHelper.GetErrorString method, then runs through a list of locales (en-US, fr-FR, de-DE, ja-JP, zh-CN). For each locale it temporarily sets the thread culture, calls GetErrorString for every standard Excel error enum, and compares the output with a predefined expected string dictionary, reporting pass/fail results for each combination.
    class Program
    {
        // Mapping of Excel error names to their expected string representations per locale
        private static readonly Dictionary<string, Dictionary<string, string>> ExpectedErrorStrings =
            new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            {
                "en-US", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "NullError", "#NULL!" },
                    { "Div0", "#DIV/0!" },
                    { "Value", "#VALUE!" },
                    { "Ref", "#REF!" },
                    { "Name", "#NAME?" },
                    { "Num", "#NUM!" },
                    { "NA", "#N/A" },
                    { "GettingData", "#GETTING_DATA" }
                }
            },
            {
                "fr-FR", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "NullError", "#NULL!" },
                    { "Div0", "#DIV/0!" },
                    { "Value", "#VALEUR!" },
                    { "Ref", "#REF!" },
                    { "Name", "#NOM?" },
                    { "Num", "#NOMBRE!" },
                    { "NA", "#N/A" },
                    { "GettingData", "#GETTING_DATA" }
                }
            },
            {
                "de-DE", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "NullError", "#NULL!" },
                    { "Div0", "#DIV/0!" },
                    { "Value", "#WERT!" },
                    { "Ref", "#BEZUG!" },
                    { "Name", "#NAME?" },
                    { "Num", "#ZAHL!" },
                    { "NA", "#NV" },
                    { "GettingData", "#GETTING_DATA" }
                }
            },
            {
                "ja-JP", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "NullError", "#NULL!" },
                    { "Div0", "#DIV/0!" },
                    { "Value", "#VALUE!" },
                    { "Ref", "#REF!" },
                    { "Name", "#NAME?" },
                    { "Num", "#NUM!" },
                    { "NA", "#N/A" },
                    { "GettingData", "#GETTING_DATA" }
                }
            },
            {
                "zh-CN", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "NullError", "#NULL!" },
                    { "Div0", "#DIV/0!" },
                    { "Value", "#VALUE!" },
                    { "Ref", "#REF!" },
                    { "Name", "#NAME?" },
                    { "Num", "#NUM!" },
                    { "NA", "#N/A" },
                    { "GettingData", "#GETTING_DATA" }
                }
            }
        };

        // List of locales to test
        private static readonly string[] Locales = { "en-US", "fr-FR", "de-DE", "ja-JP", "zh-CN" };

        static void Main()
        {
            try
            {
                // Resolve the Aspose.Cells error enum type via reflection (avoids compile‑time dependency)
                Type errorEnumType = Type.GetType("Aspose.Cells.ErrorValueType, Aspose.Cells");
                if (errorEnumType == null)
                {
                    Console.WriteLine("Unable to locate Aspose.Cells.ErrorValueType enum. Ensure Aspose.Cells assembly is referenced.");
                    return;
                }

                // Locate the GetErrorString method that accepts the enum
                MethodInfo getErrorStringMethod = typeof(CellsHelper).GetMethod("GetErrorString", new[] { errorEnumType });
                if (getErrorStringMethod == null)
                {
                    Console.WriteLine("CellsHelper.GetErrorString method not found. Ensure the Aspose.Cells version supports this API.");
                    return;
                }

                foreach (string localeName in Locales)
                {
                    CultureInfo culture = new CultureInfo(localeName);

                    if (!ExpectedErrorStrings.TryGetValue(localeName, out var expectedForLocale))
                    {
                        Console.WriteLine($"Missing expected strings for locale {localeName}");
                        continue;
                    }

                    // Set thread culture so Aspose.Cells returns locale‑specific strings
                    CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
                    Thread.CurrentThread.CurrentCulture = culture;

                    foreach (var kvp in expectedForLocale)
                    {
                        string errorName = kvp.Key;
                        string expected = kvp.Value;

                        try
                        {
                            // Parse the enum value from its name (case‑insensitive)
                            object enumValue = Enum.Parse(errorEnumType, errorName, ignoreCase: true);

                            // Invoke CellsHelper.GetErrorString via reflection
                            string actual = (string)getErrorStringMethod.Invoke(null, new[] { enumValue });

                            if (!string.Equals(actual, expected, StringComparison.Ordinal))
                            {
                                Console.WriteLine($"Locale: {localeName}, Error: {errorName} - Expected '{expected}' but got '{actual}'.");
                            }
                            else
                            {
                                Console.WriteLine($"Locale: {localeName}, Error: {errorName} - OK.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Locale: {localeName}, Error: {errorName} - Failed ({ex.Message}).");
                        }
                    }

                    // Restore original culture
                    Thread.CurrentThread.CurrentCulture = originalCulture;
                }

                Console.WriteLine("All checks completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}
