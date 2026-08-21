// Title: Reset Aspose.Cells Workbook GlobalizationSettings to Default English in C#
// Description: A C# helper that creates a fresh GlobalizationSettings instance containing the built‑in English strings (e.g., "TRUE", "FALSE", error messages) and assigns it to Workbook.Settings.GlobalizationSettings, reverting any custom localization for the entire workbook. An optional worksheetIndices parameter is kept for API compatibility but does not affect the global reset.
// Keywords: Aspose.Cells | C# | Reset GlobalizationSettings | default English globalization | SettableGlobalizationSettings | workbook culture reset | Excel localization | boolean strings TRUE FALSE | error message localization | worksheet globalization
// Common Searches: how to reset Aspose.Cells globalization to English | Aspose.Cells default English settings .NET | reset workbook culture Aspose.Cells C# | remove custom localization from Aspose.Cells workbook | globalization settings revert Aspose.Cells
// Developer Intent: Restore a workbook's globalization to the built‑in English defaults.
// Use Cases: After applying custom SettableGlobalizationSettings for a localized workbook, call the helper to ensure English boolean and error strings before saving or exporting. | When loading workbooks from unknown locales, reset globalization to guarantee consistent English output across all sheets. | In a multi‑sheet processing pipeline, use the method (with optional worksheet indices) to keep the signature stable while globally resetting culture.
// AI Prompts: Generate C# code that resets Aspose.Cells workbook globalization to default English without altering cell values. | Show how to validate that boolean cells display "TRUE" and "FALSE" after resetting GlobalizationSettings in Aspose.Cells. | Extend ResetGlobalizationSettings to accept a culture name and apply the corresponding built‑in globalization settings in C#.

using System;
using Aspose.Cells;

// A C# helper that creates a fresh GlobalizationSettings instance containing the built‑in English strings (e.g., "TRUE", "FALSE", error messages) and assigns it to Workbook.Settings.GlobalizationSettings, reverting any custom localization for the entire workbook. An optional worksheetIndices parameter is kept for API compatibility but does not affect the global reset.
public static class GlobalizationHelper
{
    /// <param name="workbook">The workbook whose globalization settings should be reset.</param>
    /// <param name="worksheetIndices">
    /// Optional array of worksheet indices that the caller is interested in.
    /// The parameter is not used internally because globalization settings are workbook‑wide,
    /// but it allows the caller to specify which sheets triggered the reset.
    /// </param>
    public static void ResetGlobalizationSettings(Workbook workbook, int[] worksheetIndices = null)
    {
        if (workbook == null)
            throw new ArgumentNullException(nameof(workbook));

        // Create a fresh GlobalizationSettings instance – this contains the default
        // English strings (e.g., "TRUE", "FALSE", error values, etc.).
        GlobalizationSettings defaultSettings = new GlobalizationSettings();

        // Assign the default settings to the workbook.
        workbook.Settings.GlobalizationSettings = defaultSettings;

        // No further per‑worksheet handling is required because the settings are
        // applied globally. The optional worksheetIndices parameter is kept for API
        // compatibility and future extensions.
    }
}

// Example usage
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue(true);
        ws.Cells["A2"].PutValue(false);
        ws.Cells["A3"].PutValue("#DIV/0!");

        // Suppose we previously applied custom globalization settings
        SettableGlobalizationSettings customSettings = new SettableGlobalizationSettings();
        customSettings.SetBooleanValueString(true, "ИСТИНА");
        customSettings.SetBooleanValueString(false, "ЛОЖЬ");
        wb.Settings.GlobalizationSettings = customSettings;

        // Now reset to default English behavior for the first worksheet (index 0)
        GlobalizationHelper.ResetGlobalizationSettings(wb, new int[] { 0 });

        // Verify that the default English strings are used
        Console.WriteLine($"Cell A1: {ws.Cells["A1"].StringValue}"); // Expected: TRUE
        Console.WriteLine($"Cell A2: {ws.Cells["A2"].StringValue}"); // Expected: FALSE
        Console.WriteLine($"Cell A3: {ws.Cells["A3"].StringValue}"); // Expected: #DIV/0!

        // Save the workbook (using the standard Aspose.Cells save method)
        wb.Save("ResetGlobalizationDemo.xlsx");
    }
}
