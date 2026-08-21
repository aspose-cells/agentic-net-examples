// Title: Unit Tests for Aspose.Cells GlobalizationSettings.GetBooleanValueString (true, false, null) Across Locales
// Description: Create comprehensive unit tests that verify Aspose.Cells .NET GlobalizationSettings.GetBooleanValueString returns the correct strings for true, false, and null values. The tests cover the default English settings, custom strings defined with SettableGlobalizationSettings, and a locale‑specific implementation (e.g., Russian).
// Keywords: Aspose.Cells | GlobalizationSettings | GetBooleanValueString | unit test | C# | localization | boolean strings | custom boolean values | Russian locale | null handling | NUnit | MSTest | xUnit
// Common Searches: Aspose.Cells unit test GetBooleanValueString null | how to test boolean localization in Aspose.Cells | custom boolean strings SettableGlobalizationSettings example | Russian boolean strings Aspose.Cells | write xUnit tests for GlobalizationSettings.GetBooleanValueString
// Developer Intent: Write unit tests that assert the correct string output for true, false, and null inputs under default, custom, and overridden globalization settings in Aspose.Cells for .NET.
// Use Cases: Verify that the default GlobalizationSettings returns "TRUE" for true and "FALSE" for false, and returns an empty string or default for null. | Confirm that SettableGlobalizationSettings returns user‑defined strings (e.g., "YES_CUSTOM", "NO_CUSTOM") for true/false and handles null consistently with the library defaults. | Ensure a derived RussianGlobalizationSettings class returns "ИСТИНА" and "ЛОЖЬ" for true/false and provides the expected result for a null argument.
// AI Prompts: Generate NUnit test methods for Aspose.Cells GlobalizationSettings.GetBooleanValueString covering true, false, and null in default, custom, and Russian locales. | Write MSTest unit tests that assert expected boolean string values and null handling when using SettableGlobalizationSettings and a derived RussianGlobalizationSettings class. | Provide xUnit test cases demonstrating how to validate GetBooleanValueString behavior for null inputs across different globalization settings in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Simple assertion helper to avoid external test frameworks
    // Create comprehensive unit tests that verify Aspose.Cells .NET GlobalizationSettings.GetBooleanValueString returns the correct strings for true, false, and null values. The tests cover the default English settings, custom strings defined with SettableGlobalizationSettings, and a locale‑specific implementation (e.g., Russian).
    internal static class SimpleAssert
    {
        public static void AreEqual(string expected, string actual, string message)
        {
            if (!string.Equals(expected, actual, StringComparison.Ordinal))
            {
                throw new Exception($"Assertion Failed: {message} Expected='{expected}', Actual='{actual}'.");
            }
        }
    }

    // Custom globalization settings for Russian locale example
    internal class RussianBooleanGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }
    }

    internal class Program
    {
        static void Main()
        {
            try
            {
                DefaultBooleanStrings_ShouldReturnUpperCaseTrueFalse();
                Console.WriteLine("DefaultBooleanStrings test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"DefaultBooleanStrings test failed: {ex.Message}");
            }

            try
            {
                CustomBooleanStrings_ShouldReturnUserDefinedValues();
                Console.WriteLine("CustomBooleanStrings test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"CustomBooleanStrings test failed: {ex.Message}");
            }

            try
            {
                LocaleSpecificBooleanStrings_ShouldReturnLocalizedValues();
                Console.WriteLine("LocaleSpecificBooleanStrings test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"LocaleSpecificBooleanStrings test failed: {ex.Message}");
            }
        }

        // Test the default globalization settings (should return "TRUE"/"FALSE")
        private static void DefaultBooleanStrings_ShouldReturnUpperCaseTrueFalse()
        {
            // Arrange
            Workbook workbook = new Workbook();
            GlobalizationSettings settings = workbook.Settings.GlobalizationSettings;

            // Act
            string trueString = settings.GetBooleanValueString(true);
            string falseString = settings.GetBooleanValueString(false);

            // Assert
            SimpleAssert.AreEqual("TRUE", trueString, "Default true string should be 'TRUE'.");
            SimpleAssert.AreEqual("FALSE", falseString, "Default false string should be 'FALSE'.");
        }

        // Test custom boolean strings set via SettableGlobalizationSettings
        private static void CustomBooleanStrings_ShouldReturnUserDefinedValues()
        {
            // Arrange
            SettableGlobalizationSettings customSettings = new SettableGlobalizationSettings();
            customSettings.SetBooleanValueString(true, "YES_CUSTOM");
            customSettings.SetBooleanValueString(false, "NO_CUSTOM");

            // Apply to a workbook (not strictly required for the method call, but mimics real usage)
            Workbook workbook = new Workbook();
            workbook.Settings.GlobalizationSettings = customSettings;

            // Act
            string trueString = customSettings.GetBooleanValueString(true);
            string falseString = customSettings.GetBooleanValueString(false);

            // Assert
            SimpleAssert.AreEqual("YES_CUSTOM", trueString, "Custom true string should match the value set via SetBooleanValueString.");
            SimpleAssert.AreEqual("NO_CUSTOM", falseString, "Custom false string should match the value set via SetBooleanValueString.");
        }

        // Test locale‑specific boolean strings by overriding GetBooleanValueString in a derived class
        private static void LocaleSpecificBooleanStrings_ShouldReturnLocalizedValues()
        {
            // Arrange
            // Russian localization example: "ИСТИНА" for true, "ЛОЖЬ" for false
            GlobalizationSettings russianSettings = new RussianBooleanGlobalizationSettings();

            Workbook workbook = new Workbook();
            workbook.Settings.GlobalizationSettings = russianSettings;

            // Act
            string trueString = russianSettings.GetBooleanValueString(true);
            string falseString = russianSettings.GetBooleanValueString(false);

            // Assert
            SimpleAssert.AreEqual("ИСТИНА", trueString, "Russian true string should be 'ИСТИНА'.");
            SimpleAssert.AreEqual("ЛОЖЬ", falseString, "Russian false string should be 'ЛОЖЬ'.");
        }
    }
}
