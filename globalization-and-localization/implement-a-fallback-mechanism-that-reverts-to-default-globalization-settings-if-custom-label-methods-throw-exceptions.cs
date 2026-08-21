// Title: Fallback to Default GlobalizationSettings in Aspose.Cells .NET When Custom Label Throws Exception
// Description: This C# example creates a workbook, applies a custom FaultyGlobalizationSettings that deliberately throws in GetAllName, catches the error, restores the built‑in GlobalizationSettings, verifies the default '(All)' label, and saves the file. It demonstrates a safe fallback pattern for globalization and localization in Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | GlobalizationSettings | fallback | exception handling | custom localization | GetAllName | label override | workbook settings | error recovery
// Common Searches: Aspose.Cells reset globalization settings after exception | How to handle GetAllName error in Aspose.Cells | C# fallback to default GlobalizationSettings | Aspose.Cells custom globalization example | Recover from faulty localization in Aspose.Cells workbook
// Developer Intent: Demonstrate how to detect a failure in a custom GlobalizationSettings implementation and automatically switch back to the built‑in default so workbook operations continue without interruption.
// Use Cases: Recover from unreliable user‑provided localization extensions during workbook generation | Guarantee that standard labels like '(All)' are always available for pivot tables | Implement defensive programming for multi‑tenant SaaS reporting services | Swap globalization settings at runtime after a runtime error | Provide robust error handling for localized Excel exports
// AI Prompts: Generate C# code that wraps Aspose.Cells GlobalizationSettings label calls in try‑catch blocks and restores the default settings on failure. | Create a unit test that verifies the fallback to GlobalizationSettings works when FaultyGlobalizationSettings throws in GetAllName. | Explain how Aspose.Cells selects the GlobalizationSettings instance at runtime and the safest way to replace it after an exception. | Suggest best practices for designing custom GlobalizationSettings classes that avoid breaking workbook processing. | Write a blog snippet describing the fallback pattern for globalization in Aspose.Cells with performance considerations.

using System;
using Aspose.Cells;

// This C# example creates a workbook, applies a custom FaultyGlobalizationSettings that deliberately throws in GetAllName, catches the error, restores the built‑in GlobalizationSettings, verifies the default '(All)' label, and saves the file. It demonstrates a safe fallback pattern for globalization and localization in Aspose.Cells.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Assign custom globalization settings that may throw exceptions
        workbook.Settings.GlobalizationSettings = new FaultyGlobalizationSettings();

        // Attempt to use a label method that could fail
        try
        {
            // This call will throw in FaultyGlobalizationSettings
            string allLabel = workbook.Settings.GlobalizationSettings.GetAllName();
            Console.WriteLine($"Custom '(All)' label: {allLabel}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Custom globalization error: {ex.Message}");

            // Revert to the default globalization settings
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings();

            // Verify that the default settings work correctly
            string defaultAll = workbook.Settings.GlobalizationSettings.GetAllName();
            Console.WriteLine($"Reverted to default '(All)' label: {defaultAll}");
        }

        // Save the workbook
        workbook.Save("FallbackGlobalizationDemo.xlsx");
    }
}

// Custom globalization that deliberately throws an exception for demonstration
class FaultyGlobalizationSettings : GlobalizationSettings
{
    // Override a label method and simulate a failure
    public override string GetAllName()
    {
        throw new InvalidOperationException("Simulated failure in GetAllName");
    }

    // Other overrides can be added as needed; this one works normally
    public override string GetColumnLabelsOfPivotTable()
    {
        return "Custom Column Labels";
    }
}
