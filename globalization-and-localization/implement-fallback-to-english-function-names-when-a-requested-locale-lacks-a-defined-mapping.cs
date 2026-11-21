// Title: Fallback to English Excel function names for missing locale mappings with Aspose.Cells in C#
// Common Searches: Aspose.Cells get localized Excel function name with fallback to English in C# | C# dictionary for Excel formula translations using Aspose.Cells | How to handle missing locale mapping for Excel functions in Aspose.Cells .NET | Insert localized formula in workbook and default to English when translation not found | Retrieve French or German Excel function name with Aspose.Cells and fallback logic
// Tags: Aspose.Cells localized formula fallback | C# Excel function name translation dictionary | Aspose.Cells locale function mapping | Excel formula internationalization .NET | fallback to English function Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example creates a Workbook, defines a dictionary of locale‑specific Excel function translations, and uses a GetFunctionName method that returns the localized name when available or falls back to the English name. It then inserts a formula with the resolved function, calculates the result, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (using the mandated create rule)
        Workbook workbook = new Workbook();

        // Example locale requested by the user
        string requestedLocale = "fr-FR";

        // The logical (English) function name we want to use
        string functionKey = "SUM";

        // Resolve the function name, falling back to English if the locale lacks a mapping
        string functionName = GetFunctionName(requestedLocale, functionKey);

        // Populate some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);

        // Insert a formula using the resolved (localized or fallback) function name
        sheet.Cells["A3"].Formula = $"{functionName}(A1:A2)";

        // Calculate the formula so the result is stored in the cell
        workbook.CalculateFormula();

        // Save the workbook (using the mandated save rule)
        workbook.Save("Result.xlsx");
    }

    // Dictionary containing known translations of function names per locale
    static readonly Dictionary<string, Dictionary<string, string>> LocaleFunctionMap = new()
    {
        // French translations
        { "fr-FR", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "SUM", "SOMME" },
                { "AVERAGE", "MOYENNE" }
            }
        },
        // German translations
        { "de-DE", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "SUM", "SUMME" },
                { "AVERAGE", "MITTELWERT" }
            }
        }
        // Additional locales can be added here
    };

    // Returns the localized function name if it exists; otherwise returns the English name
    static string GetFunctionName(string locale, string englishName)
    {
        if (LocaleFunctionMap.TryGetValue(locale, out var translations) &&
            translations.TryGetValue(englishName, out var localizedName))
        {
            return localizedName;
        }

        // Fallback to English function name
        return englishName;
    }
}
