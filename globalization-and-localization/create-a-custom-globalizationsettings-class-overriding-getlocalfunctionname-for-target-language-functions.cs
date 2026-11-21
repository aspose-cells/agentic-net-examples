// Title: Create a custom GlobalizationSettings class in C# to localize Excel function names with Aspose.Cells
// Common Searches: how to localize Excel formulas to French using Aspose.Cells C# | override GetLocalFunctionName to map English functions to other languages in Aspose.Cells | custom globalization settings example for translating SUM to SOMME in Aspose.Cells | apply custom function name translation when saving workbook with Aspose.Cells .NET | C# dictionary based function name localization with Aspose.Cells GlobalizationSettings
// Tags: custom GlobalizationSettings class Aspose.Cells | function name translation dictionary C# | localized Excel formula support Aspose.Cells | French function mapping for Excel formulas | workbook globalization settings override

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Custom globalization settings that maps English function names to localized equivalents.
// Defines a CustomGlobalizationSettings class that inherits from GlobalizationSettings, uses a case‑insensitive dictionary to translate English Excel function names (e.g., SUM, AVERAGE) to French equivalents, overrides GetLocalFunctionName to return the localized name when available, applies the settings to a Workbook, and saves the file with formulas stored in the localized form.
public class CustomGlobalizationSettings : GlobalizationSettings
{
    // Mapping of function names to their localized forms (example: French).
    private readonly Dictionary<string, string> _functionMap;

    public CustomGlobalizationSettings()
    {
        // Initialize the dictionary with desired translations.
        _functionMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "SUM", "SOMME" },
            { "AVERAGE", "MOYENNE" },
            { "MAX", "MAX" },
            { "MIN", "MIN" }
            // Add additional function mappings as needed.
        };
    }

    // Override to return the localized function name if a mapping exists.
    public override string GetLocalFunctionName(string functionName)
    {
        if (string.IsNullOrEmpty(functionName))
            return functionName;

        if (_functionMap.TryGetValue(functionName, out string localizedName))
            return localizedName;

        // Fallback to the base implementation when no custom mapping is defined.
        return base.GetLocalFunctionName(functionName);
    }
}

// Example usage of the custom globalization settings.
class Program
{
    static void Main()
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();

        // Apply the custom globalization settings to the workbook.
        workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Populate some cells.
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);

        // Use an English function name; it will be stored as the localized name defined above.
        sheet.Cells["A3"].Formula = "SUM(A1:A2)";

        // Save the workbook.
        workbook.Save("LocalizedFormula.xlsx");
    }
}
