// Title: Use Aspose.Cells for .NET to apply locale‑specific formula translations with fallback to English when a mapping is missing
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, replaces English function names with localized equivalents from a dictionary, and automatically restores the original English formula if the locale mapping is absent or calculation throws an error. | Show how to iterate through all worksheets and cells, perform case‑insensitive regex replacements of function names based on a requested locale, and implement a try‑catch fallback to English formulas using Aspose.Cells. | Demonstrate saving the workbook after processing localized formulas and ensuring calculations succeed by reverting to English when the locale does not define a translation.
// Common Searches: asp.net aspose.cells fallback to English formula when locale mapping not found | c# replace excel function names with localized versions using Aspose.Cells | how to handle missing function translations in Aspose.Cells formula calculation | iterate over worksheets and cells to apply locale specific formula translation aspose.cells | calculate workbook after translating formulas to French or German with Aspose.Cells
// Tags: Aspose.Cells formula localization fallback | C# regex function name translation Aspose.Cells | Excel workbook calculate with locale-specific functions | Aspose.Cells missing locale mapping handling | iterate worksheets cells formulas Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

// The example loads an Excel workbook, looks up a dictionary of locale‑specific function name translations, applies those translations to each formula via case‑insensitive regex, attempts to calculate the localized formula, and if calculation fails restores the original English formula before recalculating, finally saving the updated workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (using the provided load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Desired locale for function names
        string requestedLocale = "es-ES";

        // Define translations for locales that have explicit mappings
        var localeFunctionMap = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            { "fr-FR", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "SUM", "SOMME" },
                    { "AVERAGE", "MOYENNE" }
                }
            },
            { "de-DE", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "SUM", "SUMME" },
                    { "AVERAGE", "MITTELWERT" }
                }
            }
            // Additional locale mappings can be added here
        };

        // Retrieve the function translation dictionary for the requested locale, if it exists
        Dictionary<string, string> translationDict = null;
        if (!localeFunctionMap.TryGetValue(requestedLocale, out translationDict))
        {
            // No mapping for the requested locale – fallback to English (no translation needed)
            translationDict = null;
        }

        // Iterate through all worksheets and cells containing formulas
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (Cell cell in sheet.Cells)
            {
                if (cell.IsFormula)
                {
                    string originalFormula = cell.Formula;
                    string localizedFormula = originalFormula;

                    // Apply locale‑specific translations if a mapping exists
                    if (translationDict != null)
                    {
                        foreach (var kvp in translationDict)
                        {
                            // Replace English function name with its localized counterpart (case‑insensitive)
                            localizedFormula = Regex.Replace(
                                localizedFormula,
                                @"\b" + kvp.Key + @"\b",
                                kvp.Value,
                                RegexOptions.IgnoreCase);
                        }
                    }

                    // Attempt to calculate using the localized formula
                    try
                    {
                        cell.Formula = localizedFormula;
                        workbook.CalculateFormula();
                    }
                    catch
                    {
                        // If calculation fails because the function is not defined in the locale,
                        // revert to the original English formula and recalculate
                        cell.Formula = originalFormula;
                        workbook.CalculateFormula();
                    }
                }
            }
        }

        // Save the workbook (using the provided save rule)
        workbook.Save("output.xlsx");
    }
}
