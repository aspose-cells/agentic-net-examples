// Title: Detect hard‑coded numeric constants in Excel formulas and recommend named ranges using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find formulas with hard coded numbers in an Excel workbook | C# script to list Excel cells that contain numeric literals in formulas using Aspose.Cells | How to replace hard coded constants with named ranges in Excel via Aspose.Cells .NET | Detect and report constant values in Excel formulas with the Aspose.Cells API | Regex based detection of numeric constants in Excel formulas using Aspose.Cells C# example
// Tags: hard coded constant detection in Excel formulas Aspose.Cells | named range substitution for numeric literals .NET | formula scanning for constant values using Aspose.Cells C# | regex based literal identification in cell formulas | worksheet cell iteration for formula analysis Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an Excel workbook with Aspose.Cells, iterates through every worksheet and cell, and uses a regular expression to locate numeric literals that are not part of cell references within formulas. For each formula containing such literals, it prints the sheet name, cell address, the full formula, and each detected constant, then suggests replacing those constants with named ranges or parameters.
class IdentifyHardCodedConstants
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Regex to detect numeric literals that are not part of cell references
        Regex constRegex = new Regex(@"(?<![A-Z])\b\d+(\.\d+)?\b", RegexOptions.Compiled);

        // Iterate through all worksheets and cells
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            foreach (Cell cell in cells)
            {
                // Check if the cell contains a formula
                if (cell.IsFormula)
                {
                    string formula = cell.Formula;
                    MatchCollection matches = constRegex.Matches(formula);
                    if (matches.Count > 0)
                    {
                        Console.WriteLine($"Sheet: {sheet.Name}, Cell: {cell.Name}");
                        Console.WriteLine($"  Formula: {formula}");
                        foreach (Match m in matches)
                        {
                            Console.WriteLine($"  Hard‑coded constant detected: {m.Value}");
                        }
                        Console.WriteLine("  Suggestion: Replace these constants with named ranges or parameters.\n");
                    }
                }
            }
        }
    }
}
