using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    /// <summary>
    /// Simple console runner that executes the globalization settings tests.
    /// </summary>
    public static class Program
    {
        public static void Main()
        {
            RunTests();
        }

        private static void RunTests()
        {
            var tests = new GlobalizationSettingsTests();

            ExecuteTest(() => tests.GetLocalFunctionName_Sum_ReturnsLocalizedName(), "GetLocalFunctionName_Sum_ReturnsLocalizedName");
            ExecuteTest(() => tests.GetLocalFunctionName_Average_ReturnsLocalizedName(), "GetLocalFunctionName_Average_ReturnsLocalizedName");
            ExecuteTest(() => tests.GetLocalFunctionName_Max_ReturnsLocalizedName(), "GetLocalFunctionName_Max_ReturnsLocalizedName");
            ExecuteTest(() => tests.GetLocalFunctionName_Min_ReturnsLocalizedName(), "GetLocalFunctionName_Min_ReturnsLocalizedName");
            ExecuteTest(() => tests.GetLocalFunctionName_UnmappedFunction_ReturnsStandardName(), "GetLocalFunctionName_UnmappedFunction_ReturnsStandardName");
        }

        private static void ExecuteTest(Action testMethod, string testName)
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

    /// <summary>
    /// Contains tests for SettableGlobalizationSettings.
    /// </summary>
    public class GlobalizationSettingsTests
    {
        /// <summary>
        /// Helper that creates a workbook and applies a custom globalization mapping.
        /// </summary>
        private SettableGlobalizationSettings CreateSettings(string standardName, string localName)
        {
            try
            {
                // Create a new workbook (lifecycle rule)
                var workbook = new Workbook();

                // Create settings and map the standard name to the local name (bidirectional)
                var settings = new SettableGlobalizationSettings();
                settings.SetLocalFunctionName(standardName, localName, true);

                // Apply the settings to the workbook (lifecycle rule)
                workbook.Settings.GlobalizationSettings = settings;

                return settings;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to create globalization settings.", ex);
            }
        }

        public void GetLocalFunctionName_Sum_ReturnsLocalizedName()
        {
            const string standard = "SUM";
            const string local = "SOMME"; // French example

            var settings = CreateSettings(standard, local);
            string result = settings.GetLocalFunctionName(standard);

            if (result != local)
                throw new Exception($"Expected '{local}' but got '{result}'.");
        }

        public void GetLocalFunctionName_Average_ReturnsLocalizedName()
        {
            const string standard = "AVERAGE";
            const string local = "MOYENNE"; // French example

            var settings = CreateSettings(standard, local);
            string result = settings.GetLocalFunctionName(standard);

            if (result != local)
                throw new Exception($"Expected '{local}' but got '{result}'.");
        }

        public void GetLocalFunctionName_Max_ReturnsLocalizedName()
        {
            const string standard = "MAX";
            const string local = "MAXIMO"; // Spanish example

            var settings = CreateSettings(standard, local);
            string result = settings.GetLocalFunctionName(standard);

            if (result != local)
                throw new Exception($"Expected '{local}' but got '{result}'.");
        }

        public void GetLocalFunctionName_Min_ReturnsLocalizedName()
        {
            const string standard = "MIN";
            const string local = "MINIMO"; // Spanish example

            var settings = CreateSettings(standard, local);
            string result = settings.GetLocalFunctionName(standard);

            if (result != local)
                throw new Exception($"Expected '{local}' but got '{result}'.");
        }

        public void GetLocalFunctionName_UnmappedFunction_ReturnsStandardName()
        {
            // No custom mapping is defined for COUNT
            var workbook = new Workbook();
            var settings = new SettableGlobalizationSettings();
            workbook.Settings.GlobalizationSettings = settings;

            const string standard = "COUNT";

            string result = settings.GetLocalFunctionName(standard);

            if (result != standard)
                throw new Exception($"Expected standard name '{standard}' but got '{result}'.");
        }
    }
}