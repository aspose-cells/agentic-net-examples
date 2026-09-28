// Title: Write C# unit tests for BooleanStringHelper.GetBooleanString covering true, false, and null in en-US, fr-FR, and de-DE cultures
// AI Prompts: Create xUnit or NUnit test methods that set Thread.CurrentThread.CurrentUICulture to en-US, fr-FR, and de-DE, invoke BooleanStringHelper.GetBooleanString with true, false, and null, and assert the expected localized strings. | Implement a reusable helper that temporarily changes the UI culture for a test, executes a delegate, and restores the original culture after the assertion. | Build a test runner that iterates over a matrix of culture identifiers and nullable bool values, logs pass/fail outcomes, and fails the suite if any case does not match the expected result.
// Common Searches: how to unit test a method that returns localized boolean strings in .net | c# unit test for GetBooleanString with different UI cultures | testing nullable bool to string conversion for multiple locales in c# | changing currentuiculture in nunit tests for resource verification | verify french and german true/false resource strings with unit tests
// Tags: c# unit testing localized boolean strings | ui culture switching in .net tests | nullable bool resource string verification | multi‑locale unit tests for GetBooleanString | resource based boolean localization testing

using System;
using System.Globalization;
using System.Threading;

// The sample defines BooleanStringHelper.GetBooleanString, which returns a culture‑specific string for a nullable bool using Resources.BooleanTrue or Resources.BooleanFalse. A SimpleAssert helper provides basic equality checks. GetBooleanStringTests includes a method to temporarily set Thread.CurrentThread.CurrentUICulture, a collection of test cases for true, false, and null across en‑US, fr‑FR, and de‑DE, and a RunAll routine that executes each case, reports passed and failed counts, and throws on failure. Program.Main runs the suite and outputs the results.
public static class BooleanStringHelper
{
    // Returns a localized string representation of a nullable boolean.
    public static string GetBooleanString(bool? value)
    {
        if (!value.HasValue)
            return string.Empty;

        // Use the current UI culture to decide the string.
        return value.Value
            ? Resources.BooleanTrue   // e.g., "True" in en-US, "Vrai" in fr-FR
            : Resources.BooleanFalse; // e.g., "False" in en-US, "Faux" in fr-FR
    }
}

// Dummy resource class to simulate localization.
internal static class Resources
{
    public static string BooleanTrue => CultureInfo.CurrentUICulture.Name switch
    {
        "fr-FR" => "Vrai",
        "de-DE" => "Wahr",
        _ => "True"
    };

    public static string BooleanFalse => CultureInfo.CurrentUICulture.Name switch
    {
        "fr-FR" => "Faux",
        "de-DE" => "Falsch",
        _ => "False"
    };
}

// Minimal assertion helper to replace NUnit assertions.
public static class SimpleAssert
{
    public static void AreEqual(string expected, string actual)
    {
        if (!object.Equals(expected, actual))
        {
            throw new Exception($"Assert Failed: Expected '{expected}', Actual '{actual}'.");
        }
    }
}

// Test class that runs the original NUnit test cases manually.
public class GetBooleanStringTests
{
    // Executes a test action under a specific culture, restoring the original culture afterwards.
    private void ExecuteInCulture(string cultureName, Action testAction)
    {
        var originalCulture = Thread.CurrentThread.CurrentCulture;
        var originalUICulture = Thread.CurrentThread.CurrentUICulture;
        try
        {
            var culture = new CultureInfo(cultureName);
            Thread.CurrentThread.CurrentCulture = culture;
            Thread.CurrentThread.CurrentUICulture = culture;
            testAction();
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = originalCulture;
            Thread.CurrentThread.CurrentUICulture = originalUICulture;
        }
    }

    // Runs all test cases and reports results.
    public void RunAll()
    {
        var testCases = new[]
        {
            new { Culture = "en-US", Value = (bool?)true,  Expected = "True" },
            new { Culture = "en-US", Value = (bool?)false, Expected = "False" },
            new { Culture = "en-US", Value = (bool?)null,  Expected = "" },
            new { Culture = "fr-FR", Value = (bool?)true,  Expected = "Vrai" },
            new { Culture = "fr-FR", Value = (bool?)false, Expected = "Faux" },
            new { Culture = "fr-FR", Value = (bool?)null,  Expected = "" },
            new { Culture = "de-DE", Value = (bool?)true,  Expected = "Wahr" },
            new { Culture = "de-DE", Value = (bool?)false, Expected = "Falsch" },
            new { Culture = "de-DE", Value = (bool?)null,  Expected = "" }
        };

        int passed = 0;
        int failed = 0;

        foreach (var tc in testCases)
        {
            try
            {
                ExecuteInCulture(tc.Culture, () =>
                {
                    // Act
                    string result = BooleanStringHelper.GetBooleanString(tc.Value);
                    // Assert
                    SimpleAssert.AreEqual(tc.Expected, result);
                });
                passed++;
            }
            catch (Exception ex)
            {
                failed++;
                Console.WriteLine($"Test failed for Culture='{tc.Culture}', Value='{tc.Value}': {ex.Message}");
            }
        }

        Console.WriteLine($"Tests completed. Passed: {passed}, Failed: {failed}");
        if (failed > 0)
        {
            throw new Exception("One or more tests failed.");
        }
    }
}

// Entry point.
public class Program
{
    public static void Main()
    {
        try
        {
            var tests = new GetBooleanStringTests();
            tests.RunAll();
            Console.WriteLine("All tests passed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Execution terminated: {ex.Message}");
        }
    }
}
