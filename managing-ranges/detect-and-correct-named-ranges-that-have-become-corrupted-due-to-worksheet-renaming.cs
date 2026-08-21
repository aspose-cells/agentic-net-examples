// Title: Fix Corrupted Named Ranges After Worksheet Rename with Aspose.Cells for .NET
// Description: Demonstrates how to detect named ranges that point to a renamed worksheet, verify sheet existence, and automatically correct the RefersTo formula using Aspose.Cells. The sample creates a workbook, defines a range, renames the sheet, then runs a routine that extracts the sheet name via regex, substitutes a fallback sheet, and logs the changes.
// Keywords: Aspose.Cells | C# | .NET | named range correction | worksheet rename | corrupted named range | detect invalid reference | update RefersTo formula | regex sheet name extraction | Excel automation | global
// Common Searches: how to fix named ranges after sheet rename aspnet | detect invalid named range references in Aspose.Cells | update RefersTo when worksheet name changes c# | automatically correct corrupted named ranges in Excel | Aspose.Cells rename worksheet named range fix
// Developer Intent: Locate named ranges that reference non‑existent sheets and replace them with a valid worksheet name.
// Use Cases: Repair legacy workbooks where sheet names were changed after the ranges were created. | Integrate into a validation pipeline to ensure all named ranges are usable before data processing. | Provide a quick fix for user‑generated spreadsheets that contain broken range references.
// AI Prompts: Generate C# code that scans all workbook names in Aspose.Cells and updates any RefersTo formulas pointing to missing sheets. | Create a logging mechanism that records the original and corrected RefersTo strings for each fixed named range. | Rewrite the detection logic to use Workbook.Worksheets.Contains instead of indexer checks.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace NamedRangeCorrectionDemo
{
    // Demonstrates how to detect named ranges that point to a renamed worksheet, verify sheet existence, and automatically correct the RefersTo formula using Aspose.Cells. The sample creates a workbook, defines a range, renames the sheet, then runs a routine that extracts the sheet name via regex, substitutes a fallback sheet, and logs the changes.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add a worksheet with an initial name
            Workbook workbook = new Workbook();
            Worksheet oldSheet = workbook.Worksheets[0];
            oldSheet.Name = "OldSheet";

            // Populate some data in the worksheet
            oldSheet.Cells["A1"].PutValue("Item1");
            oldSheet.Cells["A2"].PutValue("Item2");
            oldSheet.Cells["A3"].PutValue("Item3");

            // Create a named range that refers to the original sheet name
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            Name namedRange = workbook.Worksheets.Names[nameIndex];
            namedRange.RefersTo = "=OldSheet!$A$1:$A$3";

            // Rename the worksheet – this makes the existing named range reference invalid
            oldSheet.Name = "NewSheet";

            // Detect and correct corrupted named ranges
            FixCorruptedNamedRanges(workbook);

            // Save the corrected workbook
            workbook.Save("CorrectedNamedRanges.xlsx");
        }

        /// <param name="wb">The workbook to process.</param>
        static void FixCorruptedNamedRanges(Workbook wb)
        {
            // Ensure there is at least one worksheet to fallback to
            if (wb.Worksheets.Count == 0) return;
            string fallbackSheetName = wb.Worksheets[0].Name;

            foreach (Name name in wb.Worksheets.Names)
            {
                string refersTo = name.RefersTo;
                if (string.IsNullOrEmpty(refersTo)) continue;

                // Extract the sheet name part from a formula like "=SheetName!$A$1:$B$2"
                Match match = Regex.Match(refersTo, @"^=([^!]+)!");
                if (!match.Success) continue; // Not a standard sheet reference

                string referencedSheet = match.Groups[1].Value;

                // Check whether the referenced sheet actually exists
                if (wb.Worksheets[referencedSheet] == null)
                {
                    // Replace the missing sheet name with the fallback sheet name
                    string correctedRefersTo = refersTo.Replace(referencedSheet, fallbackSheetName);
                    name.RefersTo = correctedRefersTo;

                    Console.WriteLine($"Updated named range '{name.Text}':");
                    Console.WriteLine($"  Old reference: {refersTo}");
                    Console.WriteLine($"  New reference: {correctedRefersTo}");
                }
            }
        }
    }
}
