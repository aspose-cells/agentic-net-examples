// Title: Check that named range RefersTo addresses stay consistent after applying German (de-DE) localization and recalculating formulas with Aspose.Cells for .NET
// AI Prompts: Load the Excel file, capture each Name object's RefersTo string, set workbook.Settings.CultureInfo to 'de-DE', call Workbook.CalculateFormula, then iterate the Names collection to compare the saved RefersTo values with the current ones and output any mismatches. | Write a C# routine that audits named‑range addresses for changes caused by applying a German locale and formula recalculation, producing a report of any altered references.
// Common Searches: aspocells ensure named ranges keep original addresses after German localization | c# detect changes in Name.RefersTo when workbook culture set to de-DE | how to test that localization does not modify named range formulas in Aspose.Cells | verify named range addresses remain unchanged after workbook.CalculateFormula with de-DE culture
// Tags: named range address stability with CultureInfo | German culture formula recalculation Aspose.Cells | C# validate Name.RefersTo after localization | Aspose.Cells workbook culture impact check

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Globalization;

// The example loads an Excel workbook, records each named range's RefersTo address, switches the workbook culture to German (de-DE), recalculates all formulas, compares the stored addresses with the current ones, logs any changes, and saves the file.
class Program
{
    static void Main()
    {
        // Load the workbook (load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Capture original named‑range references
        Dictionary<string, string> originalRefs = new Dictionary<string, string>();
        foreach (Name name in workbook.Worksheets.Names)
        {
            originalRefs[name.Text] = name.RefersTo;
        }

        // Apply German localization (de‑DE)
        workbook.Settings.CultureInfo = new CultureInfo("de-DE");

        // Recalculate formulas to reflect localization
        workbook.CalculateFormula();

        // Verify that each named‑range still points to the same reference
        foreach (Name name in workbook.Worksheets.Names)
        {
            string original = originalRefs[name.Text];
            string current = name.RefersTo;

            if (!string.Equals(original, current, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"Named range '{name.Text}' reference changed from '{original}' to '{current}'.");
            }
            else
            {
                Console.WriteLine($"Named range '{name.Text}' reference unchanged: '{current}'.");
            }
        }

        // Save the workbook (save rule)
        workbook.Save("output.xlsx");
    }
}
