// Title: Unit tests for Aspose.Cells GetErrorValueString with German and French localization
// Description: C# example that creates a custom SettableGlobalizationSettings class to map standard Excel error codes to German and French strings, provides a lightweight SimpleAssert helper, and defines comprehensive unit tests verifying default behavior, locale‑specific mappings, and workbook error display after formula calculation.
// Keywords: Aspose.Cells | GetErrorValueString | unit test | C# | globalization | localization | German error strings | French error strings | custom globalization settings | Excel error codes
// Common Searches: Aspose.Cells unit test GetErrorValueString German | How to localize Excel error messages in Aspose.Cells | Test custom globalization settings for error strings | C# Aspose.Cells error localization example | Verify localized error display in workbook
// Developer Intent: Write automated tests that confirm GetErrorValueString returns original or locale‑specific error strings and that a workbook reflects those strings when using custom globalization settings.
// Use Cases: Ensure the default SettableGlobalizationSettings returns the unchanged Excel error token for each standard error. | Validate that a LocalizedErrorGlobalizationSettings instance returns the correct German mapping for all error codes. | Validate that a LocalizedErrorGlobalizationSettings instance returns the correct French mapping for all error codes. | Confirm that assigning a custom globalization settings object to a Workbook causes cell.DisplayStringValue to show the localized error string after formula evaluation.
// AI Prompts: Generate additional unit tests for GetErrorValueString covering Spanish and Italian locales using the same pattern. | Refactor SimpleAssert into a reusable NUnit or xUnit assertion class for the Aspose.Cells test suite. | Create a test that verifies GetErrorValueString falls back to the base implementation when a mapping is missing.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Custom globalization settings that map standard error strings to locale‑specific strings.
    // C# example that creates a custom SettableGlobalizationSettings class to map standard Excel error codes to German and French strings, provides a lightweight SimpleAssert helper, and defines comprehensive unit tests verifying default behavior, locale‑specific mappings, and workbook error display after formula calculation.
    public class LocalizedErrorGlobalizationSettings : SettableGlobalizationSettings
    {
        private readonly Dictionary<string, string> _errorMap;

        public LocalizedErrorGlobalizationSettings(Dictionary<string, string> errorMap)
        {
            _errorMap = errorMap ?? throw new ArgumentNullException(nameof(errorMap));
        }

        public override string GetErrorValueString(string err)
        {
            // Return the localized string if a mapping exists; otherwise fall back to the base implementation.
            return _errorMap.TryGetValue(err, out var localized) ? localized : base.GetErrorValueString(err);
        }
    }

    // Minimal assertion helper to replace NUnit assertions.
    public static class SimpleAssert
    {
        public static void AreEqual(string expected, string actual, string message)
        {
            if (!object.Equals(expected, actual))
                throw new Exception($"Assert Failed: {message} Expected: '{expected}', Actual: '{actual}'.");
        }

        public static void IsTrue(bool condition, string message)
        {
            if (!condition)
                throw new Exception($"Assert Failed: {message}");
        }
    }

    public class GetErrorValueStringTests
    {
        // Standard Excel error codes used in the tests.
        private static readonly string[] StandardErrors =
        {
            "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#GETTING_DATA"
        };

        // Expected localized strings for German locale.
        private static readonly Dictionary<string, string> GermanErrorMap = new Dictionary<string, string>
        {
            { "#DIV/0!", "#DIV/0!" },
            { "#N/A", "#NV" },
            { "#NAME?", "#NAME?" },
            { "#NULL!", "#NULL!" },
            { "#NUM!", "#ZAHL!" },
            { "#REF!", "#BEZUG!" },
            { "#VALUE!", "#WERT!" },
            { "#GETTING_DATA", "#DATENWIRDGELEERT" }
        };

        // Expected localized strings for French locale.
        private static readonly Dictionary<string, string> FrenchErrorMap = new Dictionary<string, string>
        {
            { "#DIV/0!", "#DIV/0!" },
            { "#N/A", "#N/D" },
            { "#NAME?", "#NOM?" },
            { "#NULL!", "#VALEUR!" },
            { "#NUM!", "#NOMBRE!" },
            { "#REF!", "#REF!" },
            { "#VALUE!", "#VALEUR!" },
            { "#GETTING_DATA", "#RECUPERATION_DONNEES" }
        };

        public void DefaultSettings_ReturnsOriginalErrorString()
        {
            var settings = new SettableGlobalizationSettings();

            foreach (var err in StandardErrors)
            {
                string result = settings.GetErrorValueString(err);
                SimpleAssert.AreEqual(err, result,
                    $"Default GetErrorValueString should return the original error string for '{err}'.");
            }
        }

        public void GermanLocale_ReturnsLocalizedErrorStrings()
        {
            var settings = new LocalizedErrorGlobalizationSettings(GermanErrorMap);

            foreach (var err in StandardErrors)
            {
                string expected = GermanErrorMap[err];
                string actual = settings.GetErrorValueString(err);
                SimpleAssert.AreEqual(expected, actual,
                    $"German mapping for '{err}' is incorrect.");
            }
        }

        public void FrenchLocale_ReturnsLocalizedErrorStrings()
        {
            var settings = new LocalizedErrorGlobalizationSettings(FrenchErrorMap);

            foreach (var err in StandardErrors)
            {
                string expected = FrenchErrorMap[err];
                string actual = settings.GetErrorValueString(err);
                SimpleAssert.AreEqual(expected, actual,
                    $"French mapping for '{err}' is incorrect.");
            }
        }

        public void WorkbookUsesCustomGlobalizationSettings_ForErrorDisplay()
        {
            // Use German mapping and assign it to a workbook.
            var germanSettings = new LocalizedErrorGlobalizationSettings(GermanErrorMap);
            var workbook = new Workbook();
            workbook.Settings.GlobalizationSettings = germanSettings;
            var sheet = workbook.Worksheets[0];
            var cell = sheet.Cells["A1"];

            // Insert a formula that generates a #DIV/0! error.
            cell.Formula = "=1/0";
            workbook.CalculateFormula();

            // Verify that the cell reports an error.
            SimpleAssert.IsTrue(cell.IsErrorValue, "Cell should be flagged as an error value.");

            // The DisplayStringValue should reflect the localized string from our settings.
            string display = cell.DisplayStringValue;
            string expectedLocalized = germanSettings.GetErrorValueString("#DIV/0!");
            SimpleAssert.AreEqual(expectedLocalized, display,
                "Cell display string should use the localized error string.");
        }
    }

    class Program
    {
        static void Main()
        {
            var tests = new GetErrorValueStringTests();

            RunTest(tests.DefaultSettings_ReturnsOriginalErrorString, nameof(tests.DefaultSettings_ReturnsOriginalErrorString));
            RunTest(tests.GermanLocale_ReturnsLocalizedErrorStrings, nameof(tests.GermanLocale_ReturnsLocalizedErrorStrings));
            RunTest(tests.FrenchLocale_ReturnsLocalizedErrorStrings, nameof(tests.FrenchLocale_ReturnsLocalizedErrorStrings));
            RunTest(tests.WorkbookUsesCustomGlobalizationSettings_ForErrorDisplay,
                nameof(tests.WorkbookUsesCustomGlobalizationSettings_ForErrorDisplay));
        }

        // Executes a test method and reports success or failure.
        static void RunTest(Action testMethod, string testName)
        {
            try
            {
                testMethod();
                Console.WriteLine($"{testName}: Passed");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{testName}: Failed - {ex.Message}");
            }
        }
    }
}
