using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Custom globalization settings that return locale‑specific boolean strings.
    public class RussianBooleanGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool bv)
        {
            // Russian words for true/false.
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }
    }

    public static class SimpleAssert
    {
        public static void AreEqual(string expected, string actual, string message)
        {
            if (!string.Equals(expected, actual, StringComparison.Ordinal))
                throw new Exception($"Assert Failed: {message} Expected='{expected}', Actual='{actual}'.");
        }
    }

    public class GetBooleanValueStringTests
    {
        // Test the default implementation (no custom settings).
        public void Default_GetBooleanValueString_ReturnsEnglishStrings()
        {
            // Arrange
            Workbook workbook = new Workbook(); // default workbook
            GlobalizationSettings settings = workbook.Settings.GlobalizationSettings;

            // Act
            string trueStr = settings.GetBooleanValueString(true);
            string falseStr = settings.GetBooleanValueString(false);

            // Assert
            SimpleAssert.AreEqual("TRUE", trueStr, "Default true string should be 'TRUE'.");
            SimpleAssert.AreEqual("FALSE", falseStr, "Default false string should be 'FALSE'.");
        }

        // Test custom strings set via SettableGlobalizationSettings.
        public void SettableGlobalizationSettings_CustomBooleanStrings_AreReturned()
        {
            // Arrange
            Workbook workbook = new Workbook();
            var customSettings = new SettableGlobalizationSettings();
            customSettings.SetBooleanValueString(true, "YES_CUSTOM");
            customSettings.SetBooleanValueString(false, "NO_CUSTOM");
            workbook.Settings.GlobalizationSettings = customSettings;

            // Act
            string trueStr = customSettings.GetBooleanValueString(true);
            string falseStr = customSettings.GetBooleanValueString(false);

            // Assert
            SimpleAssert.AreEqual("YES_CUSTOM", trueStr, "Custom true string should match the value set.");
            SimpleAssert.AreEqual("NO_CUSTOM", falseStr, "Custom false string should match the value set.");
        }

        // Test a locale‑specific override (e.g., Russian) using a derived GlobalizationSettings class.
        public void CustomGlobalizationSettings_RussianLocale_ReturnsRussianStrings()
        {
            // Arrange
            Workbook workbook = new Workbook();
            workbook.Settings.GlobalizationSettings = new RussianBooleanGlobalizationSettings();

            // Act
            string trueStr = workbook.Settings.GlobalizationSettings.GetBooleanValueString(true);
            string falseStr = workbook.Settings.GlobalizationSettings.GetBooleanValueString(false);

            // Assert
            SimpleAssert.AreEqual("ИСТИНА", trueStr, "Russian true string should be 'ИСТИНА'.");
            SimpleAssert.AreEqual("ЛОЖЬ", falseStr, "Russian false string should be 'ЛОЖЬ'.");
        }

        // Test behavior when GlobalizationSettings is null (should fallback to default).
        public void Null_GlobalizationSettings_UsesDefaultImplementation()
        {
            // Arrange
            Workbook workbook = new Workbook();
            // Explicitly set to null to simulate missing settings.
            workbook.Settings.GlobalizationSettings = null;

            // Act
            // Accessing the property returns the default GlobalizationSettings instance.
            GlobalizationSettings settings = workbook.Settings.GlobalizationSettings;
            string trueStr = settings.GetBooleanValueString(true);
            string falseStr = settings.GetBooleanValueString(false);

            // Assert
            SimpleAssert.AreEqual("TRUE", trueStr, "When settings are null, true should map to 'TRUE'.");
            SimpleAssert.AreEqual("FALSE", falseStr, "When settings are null, false should map to 'FALSE'.");
        }
    }

    class Program
    {
        static void Main()
        {
            var tests = new GetBooleanValueStringTests();

            RunTest(() => tests.Default_GetBooleanValueString_ReturnsEnglishStrings(),
                nameof(tests.Default_GetBooleanValueString_ReturnsEnglishStrings));

            RunTest(() => tests.SettableGlobalizationSettings_CustomBooleanStrings_AreReturned(),
                nameof(tests.SettableGlobalizationSettings_CustomBooleanStrings_AreReturned));

            RunTest(() => tests.CustomGlobalizationSettings_RussianLocale_ReturnsRussianStrings(),
                nameof(tests.CustomGlobalizationSettings_RussianLocale_ReturnsRussianStrings));

            RunTest(() => tests.Null_GlobalizationSettings_UsesDefaultImplementation(),
                nameof(tests.Null_GlobalizationSettings_UsesDefaultImplementation));

            Console.WriteLine("All tests completed.");
        }

        // Executes a test method inside a try‑catch block and reports the result.
        private static void RunTest(Action testMethod, string testName)
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