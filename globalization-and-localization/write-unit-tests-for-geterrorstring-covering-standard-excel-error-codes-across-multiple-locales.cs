using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Custom globalization settings that map standard error codes to localized strings
    public class CustomErrorGlobalizationSettings : SettableGlobalizationSettings
    {
        private readonly Dictionary<string, string> _errorMap = new Dictionary<string, string>
        {
            { "#DIV/0!", "Division by zero" },
            { "#VALUE!", "Invalid value" },
            { "#NAME?", "Invalid name" },
            { "#N/A", "Not available" },
            { "#REF!", "Invalid reference" },
            { "#NUM!", "Invalid number" },
            { "#NULL!", "Intersection error" },
            { "#SPILL!", "Spill error" },
            { "#BUSY!", "Busy error" },
            { "#CALC!", "Calculation error" }
        };

        public override string GetErrorValueString(string err)
        {
            // Return the localized string if a mapping exists; otherwise fall back to default behavior
            return _errorMap.TryGetValue(err, out var localized) ? localized : base.GetErrorValueString(err);
        }
    }

    public static class Program
    {
        private static SettableGlobalizationSettings _settings;

        public static void Main()
        {
            try
            {
                Setup();
                RunAllTests();
                Console.WriteLine("All tests passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        // Create a workbook and assign the custom globalization settings
        private static void Setup()
        {
            try
            {
                var workbook = new Workbook();
                _settings = new CustomErrorGlobalizationSettings();
                workbook.Settings.GlobalizationSettings = _settings;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to set up workbook with custom globalization settings.", ex);
            }
        }

        // Execute each test case manually
        private static void RunAllTests()
        {
            // Standard error codes and their expected localized strings
            var testCases = new Dictionary<string, string>
            {
                { "#DIV/0!", "Division by zero" },
                { "#VALUE!", "Invalid value" },
                { "#NAME?", "Invalid name" },
                { "#N/A", "Not available" },
                { "#REF!", "Invalid reference" },
                { "#NUM!", "Invalid number" },
                { "#NULL!", "Intersection error" },
                { "#SPILL!", "Spill error" },
                { "#BUSY!", "Busy error" },
                { "#CALC!", "Calculation error" }
            };

            foreach (var kvp in testCases)
            {
                VerifyLocalizedString(kvp.Key, kvp.Value);
            }

            VerifyUnknownError();
        }

        // Verify that a known error code returns the expected localized string
        private static void VerifyLocalizedString(string errorCode, string expectedLocalized)
        {
            try
            {
                string actual = _settings.GetErrorValueString(errorCode);
                if (!string.Equals(actual, expectedLocalized, StringComparison.Ordinal))
                {
                    throw new Exception($"Error code '{errorCode}' expected '{expectedLocalized}' but got '{actual}'.");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Verification failed for error code '{errorCode}'.", ex);
            }
        }

        // Verify that an unknown error code is returned unchanged (default behavior)
        private static void VerifyUnknownError()
        {
            const string unknownError = "#UNKNOWN!";
            try
            {
                string result = _settings.GetErrorValueString(unknownError);
                if (!string.Equals(result, unknownError, StringComparison.Ordinal))
                {
                    throw new Exception($"Unknown error code expected to be unchanged but got '{result}'.");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Verification failed for unknown error code.", ex);
            }
        }
    }
}