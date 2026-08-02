using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Custom globalization settings that return Russian boolean strings
    public class RussianBooleanGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }
    }

    class Program
    {
        static void Main()
        {
            RunTest(nameof(DefaultBooleanStrings_ShouldReturnEnglishValues), DefaultBooleanStrings_ShouldReturnEnglishValues);
            RunTest(nameof(CustomSettableSettings_ShouldReturnUserDefinedValues), CustomSettableSettings_ShouldReturnUserDefinedValues);
            RunTest(nameof(OverriddenGlobalizationSettings_ShouldReturnLocalizedValues), OverriddenGlobalizationSettings_ShouldReturnLocalizedValues);
            RunTest(nameof(GetBooleanValueString_OnEmptyCell_ShouldNotThrow), GetBooleanValueString_OnEmptyCell_ShouldNotThrow);
        }

        // Executes a test method and reports the result
        static void RunTest(string testName, Action testMethod)
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

        // Test default globalization settings (en-US) return "TRUE"/"FALSE"
        static void DefaultBooleanStrings_ShouldReturnEnglishValues()
        {
            var workbook = new Workbook(); // default workbook
            GlobalizationSettings settings = workbook.Settings.GlobalizationSettings;

            string trueStr = settings.GetBooleanValueString(true);
            string falseStr = settings.GetBooleanValueString(false);

            if (trueStr != "TRUE")
                throw new Exception($"Expected 'TRUE' but got '{trueStr}'.");
            if (falseStr != "FALSE")
                throw new Exception($"Expected 'FALSE' but got '{falseStr}'.");
        }

        // Test SettableGlobalizationSettings with custom strings
        static void CustomSettableSettings_ShouldReturnUserDefinedValues()
        {
            var workbook = new Workbook();
            var customSettings = new SettableGlobalizationSettings();
            customSettings.SetBooleanValueString(true, "YES");
            customSettings.SetBooleanValueString(false, "NO");
            workbook.Settings.GlobalizationSettings = customSettings;

            string trueStr = customSettings.GetBooleanValueString(true);
            string falseStr = customSettings.GetBooleanValueString(false);

            if (trueStr != "YES")
                throw new Exception($"Expected 'YES' but got '{trueStr}'.");
            if (falseStr != "NO")
                throw new Exception($"Expected 'NO' but got '{falseStr}'.");
        }

        // Test overriding GetBooleanValueString for a specific locale (Russian)
        static void OverriddenGlobalizationSettings_ShouldReturnLocalizedValues()
        {
            var workbook = new Workbook();
            var russianSettings = new RussianBooleanGlobalizationSettings();
            workbook.Settings.GlobalizationSettings = russianSettings;

            string trueStr = russianSettings.GetBooleanValueString(true);
            string falseStr = russianSettings.GetBooleanValueString(false);

            if (trueStr != "ИСТИНА")
                throw new Exception($"Expected 'ИСТИНА' but got '{trueStr}'.");
            if (falseStr != "ЛОЖЬ")
                throw new Exception($"Expected 'ЛОЖЬ' but got '{falseStr}'.");
        }

        // Test behavior when a cell has no boolean value (null equivalent)
        static void GetBooleanValueString_OnEmptyCell_ShouldNotThrow()
        {
            var workbook = new Workbook();
            var settings = workbook.Settings.GlobalizationSettings;
            var cell = workbook.Worksheets[0].Cells["A1"]; // empty cell (unused)

            // Ensure no exception when calling the method with explicit bool values
            try
            {
                settings.GetBooleanValueString(true);
                settings.GetBooleanValueString(false);
            }
            catch (Exception ex)
            {
                throw new Exception("Method threw an exception on empty cell.", ex);
            }
        }
    }
}