// Title: Aspose.Cells C# – Fallback to English Function Names in Custom GlobalizationSettings
// Description: Demonstrates a C# CustomGlobalizationSettings class that maps German Excel functions (e.g., SUMME, MITTELWERT) to their English equivalents and automatically falls back to the original English name when no mapping exists. The example applies the settings to a workbook, evaluates formulas with both localized and fallback names, and saves the result.
// Keywords: Aspose.Cells | C# | GlobalizationSettings | function name localization | fallback to English | German Excel functions | SUMME | MITTELWERT | CalcEngineSettings | custom globalization | Excel formula translation
// Common Searches: Aspose.Cells custom GlobalizationSettings fallback English | map German Excel functions to English in Aspose.Cells | override GetLocalFunctionName for unknown locales | use localized function names with Aspose.Cells C# | Excel formula localization Aspose.Cells example | how to enable function name fallback in Aspose.Cells
// Developer Intent: Implement a GlobalizationSettings subclass that returns English function names when a locale lacks a specific mapping.
// Use Cases: Translate known German functions (SUMME, MITTELWERT) while allowing unmapped functions to run in English. | Mix localized and standard function names in the same worksheet without calculation errors. | Ensure workbook compatibility across different language settings by providing a reliable fallback mechanism.
// AI Prompts: Create a C# CustomGlobalizationSettings class for Aspose.Cells that falls back to English for any unmapped function name and show its usage in a workbook. | Extend the fallback logic to support multiple locales such as German and French while preserving English as the default. | Write unit tests for GetLocalFunctionName and GetStandardFunctionName that verify correct translation and fallback behavior.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates a C# CustomGlobalizationSettings class that maps German Excel functions (e.g., SUMME, MITTELWERT) to their English equivalents and automatically falls back to the original English name when no mapping exists. The example applies the settings to a workbook, evaluates formulas with both localized and fallback names, and saves the result.
public class CustomGlobalizationSettings : GlobalizationSettings
{
    // Mapping of standard (English) function names to their localized equivalents.
    private readonly Dictionary<string, string> _standardToLocal = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "SUM", "SUMME" },        // German example
        { "AVERAGE", "MITTELWERT" } // German example
    };

    // Return the localized name if it exists; otherwise fall back to the standard English name.
    public override string GetLocalFunctionName(string standardName)
    {
        if (_standardToLocal.TryGetValue(standardName, out var localName))
            return localName;

        // Fallback to English (standard) name when no mapping is defined.
        return standardName;
    }

    // Convert a localized name back to the standard name; fallback to the provided name if unknown.
    public override string GetStandardFunctionName(string localName)
    {
        foreach (var kvp in _standardToLocal)
        {
            if (kvp.Value.Equals(localName, StringComparison.OrdinalIgnoreCase))
                return kvp.Key;
        }

        // Fallback to the input assuming it is already a standard name.
        return localName;
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data.
            worksheet.Cells["B1"].PutValue(5);
            worksheet.Cells["B2"].PutValue(15);

            // Apply the custom globalization settings to the workbook.
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // NOTE: In recent Aspose.Cells versions, custom function names are enabled by default
            // when a custom GlobalizationSettings implementation is provided.
            // If needed, you can enable it via CalcEngineSettings (available in newer releases).

            // Formula using a localized function name that exists in the mapping.
            worksheet.Cells["B3"].Formula = "=SUMME(B1:B2)";

            // Formula using a function name that is NOT mapped; should fall back to English "SUM".
            worksheet.Cells["B4"].Formula = "=SUM(B1:B2)";

            // Calculate all formulas.
            workbook.CalculateFormula();

            // Output results to verify correct behavior.
            Console.WriteLine($"Result with localized name (SUMME): {worksheet.Cells["B3"].DoubleValue}");
            Console.WriteLine($"Result with fallback name (SUM): {worksheet.Cells["B4"].DoubleValue}");

            // Save the workbook.
            string outputPath = "FallbackLocalizationDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
