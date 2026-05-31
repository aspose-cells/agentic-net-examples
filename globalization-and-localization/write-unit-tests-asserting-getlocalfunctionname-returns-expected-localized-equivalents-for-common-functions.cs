using System;
using Aspose.Cells;
using Aspose.Cells.Settings;

namespace AsposeCellsTests
{
    class Program
    {
        static void Main()
        {
            RunTest(() => GetLocalFunctionName_ShouldReturnMappedName_ForSum(),
                nameof(GetLocalFunctionName_ShouldReturnMappedName_ForSum));
            RunTest(() => GetLocalFunctionName_ShouldReturnMappedName_ForAverage(),
                nameof(GetLocalFunctionName_ShouldReturnMappedName_ForAverage));
            RunTest(() => GetLocalFunctionName_ShouldReturnMappedName_ForMax(),
                nameof(GetLocalFunctionName_ShouldReturnMappedName_ForMax));
            RunTest(() => GetLocalFunctionName_ShouldReturnStandardName_WhenNoMappingExists(),
                nameof(GetLocalFunctionName_ShouldReturnStandardName_WhenNoMappingExists));
            RunTest(() => GetLocalFunctionName_ShouldBeCaseInsensitive_ForStandardName(),
                nameof(GetLocalFunctionName_ShouldBeCaseInsensitive_ForStandardName));
        }

        // Executes a test method and reports the result.
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

        // Helper to create a workbook with custom SettableGlobalizationSettings.
        private static Workbook CreateWorkbook(out SettableGlobalizationSettings settings)
        {
            var workbook = new Workbook();
            settings = new SettableGlobalizationSettings();
            workbook.Settings.GlobalizationSettings = settings;
            return workbook;
        }

        private static void GetLocalFunctionName_ShouldReturnMappedName_ForSum()
        {
            var workbook = CreateWorkbook(out var settings);
            settings.SetLocalFunctionName("SUM", "SOMME", true);
            string localName = settings.GetLocalFunctionName("SUM");
            Assert.AreEqual("SOMME", localName, "The localized name for SUM should be SOMME.");
        }

        private static void GetLocalFunctionName_ShouldReturnMappedName_ForAverage()
        {
            var workbook = CreateWorkbook(out var settings);
            settings.SetLocalFunctionName("AVERAGE", "MITTELWERT", true);
            string localName = settings.GetLocalFunctionName("AVERAGE");
            Assert.AreEqual("MITTELWERT", localName, "The localized name for AVERAGE should be MITTELWERT.");
        }

        private static void GetLocalFunctionName_ShouldReturnMappedName_ForMax()
        {
            var workbook = CreateWorkbook(out var settings);
            settings.SetLocalFunctionName("MAX", "MASSIMO", true);
            string localName = settings.GetLocalFunctionName("MAX");
            Assert.AreEqual("MASSIMO", localName, "The localized name for MAX should be MASSIMO.");
        }

        private static void GetLocalFunctionName_ShouldReturnStandardName_WhenNoMappingExists()
        {
            var workbook = CreateWorkbook(out var settings);
            string localName = settings.GetLocalFunctionName("MIN");
            Assert.AreEqual("MIN", localName, "Without a mapping, GetLocalFunctionName should return the standard name.");
        }

        private static void GetLocalFunctionName_ShouldBeCaseInsensitive_ForStandardName()
        {
            var workbook = CreateWorkbook(out var settings);
            settings.SetLocalFunctionName("SUM", "SUMA", true);
            string localNameUpper = settings.GetLocalFunctionName("SUM");
            string localNameLower = settings.GetLocalFunctionName("sum");
            string localNameMixed = settings.GetLocalFunctionName("SuM");
            Assert.AreEqual("SUMA", localNameUpper);
            Assert.AreEqual("SUMA", localNameLower);
            Assert.AreEqual("SUMA", localNameMixed);
        }

        // Minimal assertion helper to replace NUnit.
        static class Assert
        {
            public static void AreEqual(string expected, string actual, string message = null)
            {
                if (!string.Equals(expected, actual, StringComparison.Ordinal))
                {
                    throw new Exception(message ?? $"Expected '{expected}', but got '{actual}'.");
                }
            }
        }
    }
}