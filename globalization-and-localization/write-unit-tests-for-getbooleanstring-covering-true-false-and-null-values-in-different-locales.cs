// Title: C# unit tests for BooleanHelper.GetBooleanString covering true, false and null values across en-US, fr-FR, and de-DE cultures
// Common Searches: how to test BooleanHelper.GetBooleanString for English and French cultures in C# | unit testing localized boolean strings with CultureInfo in Aspose.Cells | C# test for GetBooleanString handling null values | verify true false string output for German culture using GetBooleanString | write custom test runner for BooleanHelper GetBooleanString method
// Tags: BooleanHelper GetBooleanString unit testing | culture-specific boolean string verification | Aspose.Cells localization test | CultureInfo true false string C# | custom C# test runner for helper methods | null handling in BooleanHelper GetBooleanString

using System;
using System.Globalization;

namespace Aspose.Cells.Tests
{
    // Helper class providing localized boolean string representations.
    // Provides a self‑contained C# test suite that validates BooleanHelper.GetBooleanString returns the correct localized strings for true, false, and null inputs in en-US, fr-FR, and de-DE cultures, using a lightweight custom test runner and simple assertions.
    public static class BooleanHelper
    {
        public static string GetBooleanString(bool? value, CultureInfo culture)
        {
            // Return empty string for null values.
            if (!value.HasValue)
                return string.Empty;

            // Determine true/false strings based on culture.
            string trueStr;
            string falseStr;

            switch (culture.Name)
            {
                case "fr-FR":
                    trueStr = "Vrai";
                    falseStr = "Faux";
                    break;
                case "de-DE":
                    trueStr = "Wahr";
                    falseStr = "Falsch";
                    break;
                default:
                    trueStr = "True";
                    falseStr = "False";
                    break;
            }

            return value.Value ? trueStr : falseStr;
        }
    }

    // Simple test runner that mimics NUnit behavior without external dependencies.
    public class GetBooleanStringTests
    {
        public void RunAll()
        {
            RunTest(GetBooleanString_True_EnUs_ReturnsTrueString);
            RunTest(GetBooleanString_False_EnUs_ReturnsFalseString);
            RunTest(GetBooleanString_Null_EnUs_ReturnsEmptyString);
            RunTest(GetBooleanString_True_FrFr_ReturnsTrueString);
            RunTest(GetBooleanString_False_FrFr_ReturnsFalseString);
            RunTest(GetBooleanString_Null_FrFr_ReturnsEmptyString);
            RunTest(GetBooleanString_True_DeDe_ReturnsTrueString);
            RunTest(GetBooleanString_False_DeDe_ReturnsFalseString);
            RunTest(GetBooleanString_Null_DeDe_ReturnsEmptyString);
        }

        private void RunTest(Action testMethod)
        {
            try
            {
                testMethod();
                Console.WriteLine($"{testMethod.Method.Name}: Passed");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{testMethod.Method.Name}: Failed - {ex.Message}");
            }
        }

        private void AssertEqual(string expected, string actual)
        {
            if (!object.Equals(expected, actual))
                throw new Exception($"Assert Failed. Expected: '{expected}', Actual: '{actual}'");
        }

        // Test true value in English (United States)
        public void GetBooleanString_True_EnUs_ReturnsTrueString()
        {
            bool? input = true;
            CultureInfo culture = new CultureInfo("en-US");
            string result = BooleanHelper.GetBooleanString(input, culture);
            AssertEqual("True", result);
        }

        // Test false value in English (United States)
        public void GetBooleanString_False_EnUs_ReturnsFalseString()
        {
            bool? input = false;
            CultureInfo culture = new CultureInfo("en-US");
            string result = BooleanHelper.GetBooleanString(input, culture);
            AssertEqual("False", result);
        }

        // Test null value in English (United States)
        public void GetBooleanString_Null_EnUs_ReturnsEmptyString()
        {
            bool? input = null;
            CultureInfo culture = new CultureInfo("en-US");
            string result = BooleanHelper.GetBooleanString(input, culture);
            AssertEqual(string.Empty, result);
        }

        // Test true value in French (France)
        public void GetBooleanString_True_FrFr_ReturnsTrueString()
        {
            bool? input = true;
            CultureInfo culture = new CultureInfo("fr-FR");
            string result = BooleanHelper.GetBooleanString(input, culture);
            AssertEqual("Vrai", result);
        }

        // Test false value in French (France)
        public void GetBooleanString_False_FrFr_ReturnsFalseString()
        {
            bool? input = false;
            CultureInfo culture = new CultureInfo("fr-FR");
            string result = BooleanHelper.GetBooleanString(input, culture);
            AssertEqual("Faux", result);
        }

        // Test null value in French (France)
        public void GetBooleanString_Null_FrFr_ReturnsEmptyString()
        {
            bool? input = null;
            CultureInfo culture = new CultureInfo("fr-FR");
            string result = BooleanHelper.GetBooleanString(input, culture);
            AssertEqual(string.Empty, result);
        }

        // Test true value in German (Germany)
        public void GetBooleanString_True_DeDe_ReturnsTrueString()
        {
            bool? input = true;
            CultureInfo culture = new CultureInfo("de-DE");
            string result = BooleanHelper.GetBooleanString(input, culture);
            AssertEqual("Wahr", result);
        }

        // Test false value in German (Germany)
        public void GetBooleanString_False_DeDe_ReturnsFalseString()
        {
            bool? input = false;
            CultureInfo culture = new CultureInfo("de-DE");
            string result = BooleanHelper.GetBooleanString(input, culture);
            AssertEqual("Falsch", result);
        }

        // Test null value in German (Germany)
        public void GetBooleanString_Null_DeDe_ReturnsEmptyString()
        {
            bool? input = null;
            CultureInfo culture = new CultureInfo("de-DE");
            string result = BooleanHelper.GetBooleanString(input, culture);
            AssertEqual(string.Empty, result);
        }
    }

    // Entry point.
    class Program
    {
        static void Main()
        {
            try
            {
                var tests = new GetBooleanStringTests();
                tests.RunAll();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
