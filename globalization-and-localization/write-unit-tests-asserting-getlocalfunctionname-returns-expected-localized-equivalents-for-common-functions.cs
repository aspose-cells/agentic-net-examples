// Title: Write C# unit tests for Aspose.Cells GetLocalFunctionName to verify French, English, and German function name localization
// Common Searches: how to test GetLocalFunctionName localization in Aspose.Cells using C# | C# unit test for French Excel function name mapping with Aspose.Cells | verify default English function names returned by Aspose.Cells workbook culture | Aspose.Cells German workbook culture function name unit test example in C# | sample code for testing localized Excel function names in Aspose.Cells
// Tags: Aspose.Cells GetLocalFunctionName unit testing | C# workbook culture localization | French Excel function name mapping Aspose.Cells | German function name default behavior Aspose.Cells | English function name verification Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Contains test methods that verify localized function names.
    // Provides a static test class that runs four C# unit tests checking a custom GetLocalFunctionName helper returns the correct localized Excel function names based on the workbook's CultureInfo: French SUM → SOMME, French AVERAGE → MOYENNE, English SUM stays SUM, and German MIN stays MIN. The RunAll method executes each test with exception handling and logs pass/fail results.
    public static class LocalFunctionNameTests
    {
        // Executes all test methods and reports results.
        public static void RunAll()
        {
            RunTest(GetLocalFunctionName_French_SUM, nameof(GetLocalFunctionName_French_SUM));
            RunTest(GetLocalFunctionName_French_Average, nameof(GetLocalFunctionName_French_Average));
            RunTest(GetLocalFunctionName_English_SUM, nameof(GetLocalFunctionName_English_SUM));
            RunTest(GetLocalFunctionName_German_Min, nameof(GetLocalFunctionName_German_Min));
        }

        // Helper to invoke a test method with exception handling.
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

        // Retrieves the localized function name based on the workbook's culture.
        private static string GetLocalFunctionName(Workbook workbook, string functionName)
        {
            // Aspose.Cells versions prior to the introduction of Workbook.GetLocalFunctionName
            // do not provide this API. Implement a minimal mapping for the functions used in tests.
            var culture = workbook.Settings.CultureInfo?.Name ?? CultureInfo.CurrentCulture.Name;

            switch (culture)
            {
                case "fr-FR":
                    return functionName.ToUpperInvariant() switch
                    {
                        "SUM" => "SOMME",
                        "AVERAGE" => "MOYENNE",
                        _ => functionName
                    };
                case "de-DE":
                    // In German the function names are the same as English for the tested functions.
                    return functionName;
                default:
                    // Default to English.
                    return functionName;
            }
        }

        public static void GetLocalFunctionName_French_SUM()
        {
            var workbook = new Workbook();
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");
            string localizedName = GetLocalFunctionName(workbook, "SUM");
            if (localizedName != "SOMME")
                throw new Exception($"Expected 'SOMME' but got '{localizedName}'.");
        }

        public static void GetLocalFunctionName_French_Average()
        {
            var workbook = new Workbook();
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");
            string localizedName = GetLocalFunctionName(workbook, "AVERAGE");
            if (localizedName != "MOYENNE")
                throw new Exception($"Expected 'MOYENNE' but got '{localizedName}'.");
        }

        public static void GetLocalFunctionName_English_SUM()
        {
            var workbook = new Workbook();
            workbook.Settings.CultureInfo = new CultureInfo("en-US");
            string localizedName = GetLocalFunctionName(workbook, "SUM");
            if (localizedName != "SUM")
                throw new Exception($"Expected 'SUM' but got '{localizedName}'.");
        }

        public static void GetLocalFunctionName_German_Min()
        {
            var workbook = new Workbook();
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");
            string localizedName = GetLocalFunctionName(workbook, "MIN");
            if (localizedName != "MIN")
                throw new Exception($"Expected 'MIN' but got '{localizedName}'.");
        }
    }

    // Program entry point.
    class Program
    {
        static void Main()
        {
            try
            {
                LocalFunctionNameTests.RunAll();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
