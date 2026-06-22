using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace NamedRangeCorrectionDemo
{
    class Program
    {
        static void Main()
        {
            // Load the workbook that may contain corrupted named ranges
            Workbook workbook = new Workbook("input.xlsx");

            // Build a set of current worksheet names for quick lookup
            HashSet<string> sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Worksheet ws in workbook.Worksheets)
            {
                sheetNames.Add(ws.Name);
            }

            // OPTIONAL: Mapping of old sheet names to new ones (if known)
            // In real scenarios this could be built from a log of rename operations
            Dictionary<string, string> renameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // {"OldSheetName", "NewSheetName"},
                // Example:
                // {"Data_2020", "Data_2021"}
            };

            // Regular expression to capture sheet name part before '!' in a RefersTo formula
            // Handles cases like =Sheet1!$A$1:$B$2 or ='My Sheet'!$A$1
            Regex sheetRegex = new Regex(@"^=?'?(?<sheet>[^'!]+)'?!", RegexOptions.Compiled);

            // Iterate through all defined names in the workbook
            foreach (Name definedName in workbook.Worksheets.Names)
            {
                string refersTo = definedName.RefersTo; // e.g., "=Sheet1!$A$1:$A$10"
                if (string.IsNullOrEmpty(refersTo))
                    continue;

                // Extract the sheet name from the RefersTo string
                Match match = sheetRegex.Match(refersTo);
                if (!match.Success)
                    continue; // Unable to parse, skip

                string originalSheet = match.Groups["sheet"].Value;

                // Check if the sheet still exists
                if (sheetNames.Contains(originalSheet))
                    continue; // Reference is valid

                // Try to find a replacement sheet name
                string newSheet = null;

                // 1) Use explicit rename map if provided
                if (renameMap.TryGetValue(originalSheet, out string mapped))
                {
                    newSheet = mapped;
                }
                else
                {
                    // 2) Fallback: look for a sheet with a similar name (case‑insensitive contains)
                    foreach (string existing in sheetNames)
                    {
                        if (existing.IndexOf(originalSheet, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            newSheet = existing;
                            break;
                        }
                    }
                }

                if (newSheet == null)
                {
                    // No suitable replacement found – optionally log or leave as is
                    Console.WriteLine($"Unable to resolve sheet '{originalSheet}' for named range '{definedName.Text}'.");
                    continue;
                }

                // Build the corrected RefersTo string by replacing the old sheet name
                // Preserve any leading '=' and possible surrounding quotes
                string correctedRefersTo = Regex.Replace(
                    refersTo,
                    @"^=?'?[^'!]+?'?!",
                    m => (refersTo.StartsWith("=") ? "=" : "") + (refersTo.Contains("'") ? $"'{newSheet}'!" : $"{newSheet}!"),
                    RegexOptions.Compiled);

                // Apply the corrected reference
                definedName.RefersTo = correctedRefersTo;

                Console.WriteLine($"Corrected named range '{definedName.Text}': '{refersTo}' => '{correctedRefersTo}'");
            }

            // Save the corrected workbook
            workbook.Save("output_corrected.xlsx");
        }
    }
}