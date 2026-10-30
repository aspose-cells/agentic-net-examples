// Title: C# – Merge two Excel workbooks with Aspose.Cells while preserving named ranges and handling sheet name conflicts
// Common Searches: C# Aspose.Cells merge two workbooks and keep each workbook's named ranges | How to copy worksheets from one Excel file to another without overwriting existing sheet names using Aspose.Cells | Adjust named range RefersTo after renaming sheets when merging Excel files in .NET | Resolve duplicate worksheet names when combining Excel workbooks with Aspose.Cells | Merge workbook‑level named ranges from two Excel files in C#
// Tags: merge workbooks preserving named ranges Aspose.Cells | copy worksheets duplicate name handling Aspose.Cells | adjust RefersTo formulas after sheet rename Aspose.Cells | C# workbook merge conflict resolution Aspose.Cells | Excel file merge named range mapping .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads two Excel workbooks, copies each worksheet from the second workbook into the first while detecting and renaming any conflicting sheet names, records a mapping of original to new sheet names, optionally merges workbook‑level named ranges by updating their RefersTo formulas using the mapping, and saves the combined workbook. Helper methods provide case‑insensitive replacement to ensure RefersTo strings reference the renamed sheets correctly.
class WorkbookMerger
{
    static void Main()
    {
        // Paths to source workbooks and the output workbook
        string path1 = "Workbook1.xlsx";
        string path2 = "Workbook2.xlsx";
        string outputPath = "MergedWorkbook.xlsx";

        try
        {
            // Verify that source files exist
            if (!File.Exists(path1))
                throw new FileNotFoundException($"Source file not found: {path1}");
            if (!File.Exists(path2))
                throw new FileNotFoundException($"Source file not found: {path2}");

            // Load the two workbooks
            Workbook wb1 = new Workbook(path1);
            Workbook wb2 = new Workbook(path2);

            // Map original sheet names (from wb2) to the new sheet names after copy
            Dictionary<string, string> sheetNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // -------------------------------------------------
            // 1. Copy worksheets from wb2 into wb1, handling name conflicts
            // -------------------------------------------------
            foreach (Worksheet srcSheet in wb2.Worksheets)
            {
                string targetName = srcSheet.Name;

                // Resolve sheet name conflict with wb1
                if (wb1.Worksheets[targetName] != null)
                {
                    int suffix = 1;
                    string baseName = targetName;
                    while (wb1.Worksheets[$"{baseName}_Copy{suffix}"] != null)
                        suffix++;
                    targetName = $"{baseName}_Copy{suffix}";
                }

                // Copy the worksheet (AddCopy returns the new sheet index)
                int newIndex = wb1.Worksheets.AddCopy(srcSheet.Name);
                Worksheet copiedSheet = wb1.Worksheets[newIndex];
                copiedSheet.Name = targetName;

                // Record the mapping for later named‑range adjustment (if needed)
                sheetNameMap[srcSheet.Name] = targetName;
            }

            // -------------------------------------------------
            // 2. (Optional) Merge workbook‑level named ranges
            // -------------------------------------------------
            /*
            foreach (Name srcName in wb2.Names)
            {
                string newName = srcName.Name;
                if (wb1.Names.Contains(newName))
                {
                    int suffix = 1;
                    string baseName = newName;
                    while (wb1.Names.Contains($"{baseName}_2"))
                        suffix++;
                    newName = $"{baseName}_2";
                }

                string adjustedRefersTo = AdjustRefersTo(srcName.RefersTo, sheetNameMap);
                Name mergedName = wb1.Names.Add(newName);
                mergedName.RefersTo = adjustedRefersTo;
            }
            */

            // -------------------------------------------------
            // 3. Save the merged workbook
            // -------------------------------------------------
            wb1.Save(outputPath);
            Console.WriteLine($"Merged workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Adjusts RefersTo strings to use the new sheet names after copying
    private static string AdjustRefersTo(string refersTo, Dictionary<string, string> sheetMap)
    {
        if (string.IsNullOrEmpty(refersTo))
            return refersTo;

        string result = refersTo;

        foreach (var kvp in sheetMap)
        {
            // Replace occurrences of the original sheet name followed by '!' (e.g., Sheet1!)
            string oldPattern = $"{kvp.Key}!";
            string newPattern = $"{kvp.Value}!";

            result = ReplaceIgnoreCase(result, oldPattern, newPattern);
        }

        return result;
    }

    // Case‑insensitive string replacement
    private static string ReplaceIgnoreCase(string input, string search, string replacement)
    {
        return Regex.Replace(
            input,
            Regex.Escape(search),
            replacement.Replace("$", "$$"),
            RegexOptions.IgnoreCase);
    }
}
