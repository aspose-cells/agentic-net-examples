// Title: Convert VLOOKUP formulas to INDEX‑MATCH in an Excel workbook using Aspose.Cells for .NET
// Common Searches: c# Aspose.Cells replace VLOOKUP with INDEX MATCH in an existing workbook | programmatically change VLOOKUP formulas to INDEX-MATCH using Aspose.Cells | batch convert VLOOKUP to INDEX MATCH in Excel files with C# | how to update Excel formulas from VLOOKUP to INDEX-MATCH with Aspose.Cells | regex based VLOOKUP to INDEX-MATCH transformation in C#
// Tags: replace VLOOKUP with INDEX-MATCH Aspose.Cells | batch formula transformation C# Excel | regex formula conversion Aspose.Cells | optimize lookup performance Excel C# | programmatic workbook formula update | Aspose.Cells formula manipulation

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet's used range, detects cells containing VLOOKUP formulas, uses a regular expression to extract the arguments, and replaces each VLOOKUP with an equivalent INDEX‑MATCH expression before saving the updated file.
class VlookupToIndexMatchConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath); // {LoadWorkbook}

            // Regex to capture VLOOKUP arguments
            Regex vlookupRegex = new Regex(
                @"VLOOKUP\s*\(\s*(?<lookup>[^,]+)\s*,\s*(?<table>[^,]+)\s*,\s*(?<col>\d+)\s*(?:,\s*(?<range>[^)]+))?\s*\)",
                RegexOptions.IgnoreCase);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range (as Aspose.Cells.Range object)
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                if (usedRange == null || usedRange.RowCount == 0 || usedRange.ColumnCount == 0)
                    continue; // Skip empty sheets

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Iterate through each cell in the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Process only cells that contain a VLOOKUP formula
                        if (cell.IsFormula && cell.Formula.Contains("VLOOKUP", StringComparison.OrdinalIgnoreCase))
                        {
                            string originalFormula = cell.Formula;

                            // Replace VLOOKUP with INDEX-MATCH
                            string newFormula = vlookupRegex.Replace(originalFormula, match =>
                            {
                                string lookup = match.Groups["lookup"].Value.Trim();
                                string table = match.Groups["table"].Value.Trim();
                                string colIndex = match.Groups["col"].Value.Trim();

                                // Construct INDEX-MATCH formula (exact match)
                                return $"INDEX({table}, MATCH({lookup}, INDEX({table},0,1), 0), {colIndex})";
                            });

                            cell.Formula = newFormula;
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath); // {SaveWorkbook}
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
