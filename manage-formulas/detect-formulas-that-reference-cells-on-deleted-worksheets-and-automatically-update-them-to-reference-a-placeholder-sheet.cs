// Title: How to detect and replace formulas that reference deleted worksheets with a placeholder sheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find formulas that reference non‑existent worksheets | Replace missing sheet references in Excel formulas with a placeholder using Aspose.Cells | C# code to update formulas after deleting a worksheet in an Excel file | How to automatically change sheet name in formulas when the sheet is removed with Aspose.Cells | Detect invalid worksheet references in a workbook and redirect them to a dummy sheet in .NET
// Tags: detect missing worksheet references in Excel formulas using Aspose.Cells | swap invalid sheet reference to placeholder worksheet Aspose.Cells | update formula sheet names after worksheet deletion C# | regex based sheet reference extraction Aspose.Cells | ensure placeholder sheet exists before formula correction .NET

using Aspose.Cells;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

// The example loads an Excel workbook, adds a "Placeholder" worksheet if it does not exist, scans every formula for sheet references using a regular expression, and rewrites any reference that points to a deleted worksheet to point to the placeholder sheet before saving the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string placeholderName = "Placeholder";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load workbook (load rule)
            Workbook workbook = new Workbook(inputPath);

            // Ensure placeholder sheet exists
            if (!workbook.Worksheets.Any(ws => ws.Name.Equals(placeholderName, StringComparison.OrdinalIgnoreCase)))
            {
                workbook.Worksheets.Add(placeholderName);
            }

            // Regex to locate sheet references in formulas (e.g., Sheet1!A1 or 'Sheet 1'!A1)
            Regex sheetRefRegex = new Regex(@"('([^']+)'|([^'!]+))!", RegexOptions.Compiled);

            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Skip the placeholder sheet itself
                if (ws.Name.Equals(placeholderName, StringComparison.OrdinalIgnoreCase))
                    continue;

                // Iterate through all cells that contain formulas
                foreach (Cell cell in ws.Cells)
                {
                    if (!cell.IsFormula)
                        continue;

                    string formula = cell.Formula;
                    var matches = sheetRefRegex.Matches(formula);
                    bool updated = false;

                    foreach (Match match in matches)
                    {
                        // Extract the sheet name from the match
                        string sheetName = match.Groups[2].Success ? match.Groups[2].Value : match.Groups[3].Value;

                        // If the referenced sheet no longer exists, replace it with the placeholder
                        if (!workbook.Worksheets.Any(s => s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
                        {
                            string oldRef = match.Value;               // e.g., Sheet1! or 'Sheet 1'!
                            string newRef = $"'{placeholderName}'!";   // placeholder reference
                            formula = formula.Replace(oldRef, newRef);
                            updated = true;
                        }
                    }

                    // Apply the updated formula back to the cell
                    if (updated)
                    {
                        cell.Formula = formula;
                    }
                }
            }

            // Save workbook (save rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
