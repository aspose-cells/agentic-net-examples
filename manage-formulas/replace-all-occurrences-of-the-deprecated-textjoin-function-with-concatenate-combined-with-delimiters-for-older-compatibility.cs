// Title: Convert TEXTJOIN to CONCATENATE in Excel using Aspose.Cells for .NET
// Description: A C# example that loads an Excel file, scans every worksheet, detects TEXTJOIN formulas, rewrites them as CONCATENATE expressions with the original delimiter, and saves the workbook—ensuring compatibility with legacy Excel versions (2007‑2010).
// Keywords: Aspose.Cells | C# Excel formula conversion | TEXTJOIN replacement | CONCATENATE fallback | legacy Excel compatibility | server‑side spreadsheet processing | Excel 2007 support | global spreadsheet automation | US developers | EU data handling
// Common Searches: how to replace TEXTJOIN with CONCATENATE using Aspose.Cells | C# code to convert Excel TEXTJOIN formulas | update old Excel workbooks for 2007 compatibility | Aspose.Cells replace deprecated functions | bulk formula conversion .NET
// Developer Intent: Rewrite all TEXTJOIN formulas in a workbook to equivalent CONCATENATE formulas for older Excel versions.
// Use Cases: Prepare a workbook for distribution to users running Excel 2007 or earlier. | Automate bulk migration of a spreadsheet library to legacy‑compatible formulas on a server. | Integrate formula conversion into a CI pipeline that validates Excel files before release.
// AI Prompts: Generate C# code with Aspose.Cells that finds every TEXTJOIN formula and replaces it with CONCATENATE while preserving delimiters and the ignore‑empty flag. | Create unit tests for ConvertTextJoinToConcat covering quoted delimiters, multiple arguments, and formulas without TEXTJOIN. | Explain how to extend the conversion logic to handle range arguments such as TEXTJOIN(",",TRUE,A1:A5).

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A C# example that loads an Excel file, scans every worksheet, detects TEXTJOIN formulas, rewrites them as CONCATENATE expressions with the original delimiter, and saves the workbook—ensuring compatibility with legacy Excel versions (2007‑2010).
    public class TextJoinReplacementDemo
    {
        // Converts a TEXTJOIN formula to an equivalent CONCATENATE formula.
        // Supports simple TEXTJOIN usage: TEXTJOIN(delimiter, ignore_empty, arg1, arg2, ...)
        private static string ConvertTextJoinToConcat(string formula)
        {
            // Remove leading '=' if present for easier processing
            string cleanFormula = formula.StartsWith("=") ? formula.Substring(1) : formula;

            // Regex to capture delimiter, ignore_empty flag, and the rest of the arguments
            var match = Regex.Match(
                cleanFormula,
                @"TEXTJOIN\s*\(\s*(?<delim>[^,]+)\s*,\s*(?<ignore>[^,]+)\s*,\s*(?<args>.+)\)",
                RegexOptions.IgnoreCase);

            if (!match.Success)
                return formula; // Return original if pattern not matched

            // Extract delimiter (remove surrounding quotes if any)
            string delimiter = match.Groups["delim"].Value.Trim().Trim('\"', '\'');

            // Extract arguments string and split by commas not inside quotes
            string argsPart = match.Groups["args"].Value;
            var argList = new List<string>();
            int start = 0;
            bool inQuote = false;
            for (int i = 0; i < argsPart.Length; i++)
            {
                if (argsPart[i] == '\"')
                    inQuote = !inQuote;
                else if (argsPart[i] == ',' && !inQuote)
                {
                    argList.Add(argsPart.Substring(start, i - start).Trim());
                    start = i + 1;
                }
            }
            // Add last argument
            argList.Add(argsPart.Substring(start).Trim());

            // Build CONCATENATE formula by interleaving delimiter between arguments
            var concatBuilder = new StringBuilder();
            concatBuilder.Append("CONCATENATE(");
            for (int i = 0; i < argList.Count; i++)
            {
                concatBuilder.Append(argList[i]);
                if (i < argList.Count - 1)
                {
                    concatBuilder.Append($", \"{delimiter}\", ");
                }
            }
            concatBuilder.Append(")");

            // Preserve leading '=' if it existed
            return "=" + concatBuilder.ToString();
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load workbook if file exists; otherwise create a new empty workbook
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Iterate through all worksheets and cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell != null && cell.IsFormula && cell.Formula.Contains("TEXTJOIN", StringComparison.OrdinalIgnoreCase))
                        {
                            string originalFormula = cell.Formula;
                            string newFormula = ConvertTextJoinToConcat(originalFormula);
                            cell.Formula = newFormula;
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }

        // Entry point
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
