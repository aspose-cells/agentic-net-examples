// Title: C# Unit Tests for Aspose.Cells GetLocalFunctionName – Verify French Function Name Mappings
// Description: Demonstrates how to configure SettableGlobalizationSettings with French equivalents for SUM, AVERAGE and MAX, then validates GetLocalFunctionName returns those localized names or falls back to the original English name when no mapping exists. Includes a simple test runner and custom assertion helper.
// Keywords: Aspose.Cells | GetLocalFunctionName | SettableGlobalizationSettings | C# unit test | function name localization | French Excel functions | SUM SOMME | AVERAGE MOYENNE | MAX MAXIMUM | unmapped function handling
// Common Searches: Aspose.Cells test GetLocalFunctionName | C# unit test for localized Excel functions | SettableGlobalizationSettings French mapping example | how to verify function name localization in Aspose.Cells | GetLocalFunctionName returns original name when not mapped
// Developer Intent: Create automated tests that confirm GetLocalFunctionName returns the correct French translation for mapped functions and the default English name for unmapped functions in Aspose.Cells.
// Use Cases: Ensure French workbook users see SUM as SOMME, AVERAGE as MOYENNE, and MAX as MAXIMUM. | Detect missing localization entries by checking that unmapped functions like MIN return their English identifiers. | Integrate the test suite into CI pipelines to guard against regression in globalization settings.
// AI Prompts: Generate NUnit test methods that assert GetLocalFunctionName returns SOMME, MOYENNE, and MAXIMUM for SUM, AVERAGE, and MAX, and returns MIN unchanged when not mapped. | Rewrite the provided test class using xUnit with [Fact] attributes and Assert.Equal, preserving the French mapping logic. | Create a reusable logging assertion for Aspose.Cells localization tests that captures expected vs. actual function names.

using System;
using Aspose.Cells;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    // Demonstrates how to configure SettableGlobalizationSettings with French equivalents for SUM, AVERAGE and MAX, then validates GetLocalFunctionName returns those localized names or falls back to the original English name when no mapping exists. Includes a simple test runner and custom assertion helper.
    public class GlobalizationSettingsTests
    {
        // Helper method to create a workbook with custom function name mappings
        private Workbook CreateWorkbookWithMappings()
        {
            var workbook = new Workbook();

            // Use SettableGlobalizationSettings to define local equivalents
            var settings = new SettableGlobalizationSettings();
            settings.SetLocalFunctionName("SUM", "SOMME", true);        // French for SUM
            settings.SetLocalFunctionName("AVERAGE", "MOYENNE", true); // French for AVERAGE
            settings.SetLocalFunctionName("MAX", "MAXIMUM", true);    // Example for MAX

            // Apply the settings to the workbook
            workbook.Settings.GlobalizationSettings = settings;

            return workbook;
        }

        // Simple assertion helper
        private void AssertEqual(string expected, string actual, string message)
        {
            if (!string.Equals(expected, actual, StringComparison.Ordinal))
            {
                throw new Exception($"{message} Expected: '{expected}', Actual: '{actual}'.");
            }
        }

        public void GetLocalFunctionName_Sum_ReturnsSOMME()
        {
            var workbook = CreateWorkbookWithMappings();
            var settings = (SettableGlobalizationSettings)workbook.Settings.GlobalizationSettings;

            string localName = settings.GetLocalFunctionName("SUM");

            AssertEqual("SOMME", localName, "SUM mapping failed");
        }

        public void GetLocalFunctionName_Average_ReturnsMOYENNE()
        {
            var workbook = CreateWorkbookWithMappings();
            var settings = (SettableGlobalizationSettings)workbook.Settings.GlobalizationSettings;

            string localName = settings.GetLocalFunctionName("AVERAGE");

            AssertEqual("MOYENNE", localName, "AVERAGE mapping failed");
        }

        public void GetLocalFunctionName_Max_ReturnsMAXIMUM()
        {
            var workbook = CreateWorkbookWithMappings();
            var settings = (SettableGlobalizationSettings)workbook.Settings.GlobalizationSettings;

            string localName = settings.GetLocalFunctionName("MAX");

            AssertEqual("MAXIMUM", localName, "MAX mapping failed");
        }

        public void GetLocalFunctionName_UnmappedFunction_ReturnsStandardName()
        {
            var workbook = CreateWorkbookWithMappings();
            var settings = (SettableGlobalizationSettings)workbook.Settings.GlobalizationSettings;

            // "MIN" was not mapped, so the method should return the original standard name
            string localName = settings.GetLocalFunctionName("MIN");

            AssertEqual("MIN", localName, "Unmapped function handling failed");
        }

        // Entry point to run the tests
        public static void Main()
        {
            var tests = new GlobalizationSettingsTests();

            RunTest(tests.GetLocalFunctionName_Sum_ReturnsSOMME, "GetLocalFunctionName_Sum_ReturnsSOMME");
            RunTest(tests.GetLocalFunctionName_Average_ReturnsMOYENNE, "GetLocalFunctionName_Average_ReturnsMOYENNE");
            RunTest(tests.GetLocalFunctionName_Max_ReturnsMAXIMUM, "GetLocalFunctionName_Max_ReturnsMAXIMUM");
            RunTest(tests.GetLocalFunctionName_UnmappedFunction_ReturnsStandardName, "GetLocalFunctionName_UnmappedFunction_ReturnsStandardName");
        }

        // Helper to execute a test method with exception handling
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
