// Title: Convert TEXTJOIN to CONCATENATE in Excel files with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, scans all worksheets for TEXTJOIN formulas, extracts the delimiter and source arguments, and rewrites each formula as an equivalent CONCATENATE expression that preserves the original delimiter. Handles both range references and comma‑separated lists, then saves the updated file.
// Keywords: Aspose.Cells C# replace TEXTJOIN | convert TEXTJOIN to CONCATENATE | Excel formula migration .NET | programmatic formula edit | legacy Excel compatibility | C# regex formula replace | Aspose.Cells formula manipulation
// Common Searches: C# Aspose.Cells replace TEXTJOIN | How to change TEXTJOIN to CONCATENATE in .NET | Batch convert Excel TEXTJOIN formulas | Programmatically edit Excel formulas with Aspose.Cells | Remove deprecated TEXTJOIN function
// Developer Intent: Replace every TEXTJOIN function in a workbook with an equivalent CONCATENATE expression for older Excel versions.
// Use Cases: Modernize legacy spreadsheets before opening them in Excel 2010 or earlier. | Automate bulk conversion of customer‑uploaded workbooks during data‑import pipelines. | Prepare files for environments that lack TEXTJOIN support, such as older Office suites or third‑party parsers.
// AI Prompts: Generate C# code using Aspose.Cells that finds all TEXTJOIN formulas in a workbook and rewrites them as CONCATENATE formulas, preserving delimiters and handling both range and list sources. | Show how to add logging to the ReplaceTextJoinWithConcatenate example so each converted formula is recorded, and skip conversion when the ignore_empty argument is TRUE. | Write unit tests for the conversion logic that verify correct output for TEXTJOIN with a range, with a list of cells, and with mixed arguments.

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, scans all worksheets for TEXTJOIN formulas, extracts the delimiter and source arguments, and rewrites each formula as an equivalent CONCATENATE expression that preserves the original delimiter. Handles both range references and comma‑separated lists, then saves the updated file.
    public class ReplaceTextJoinWithConcatenate
    {
        public static void Main()
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

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Regex to capture TEXTJOIN parameters: TEXTJOIN(delimiter, ignore_empty, source)
            Regex textJoinRegex = new Regex(
                @"TEXTJOIN\(\s*(?<delim>[^,]+)\s*,\s*(?<ignore>[^,]+)\s*,\s*(?<source>.+?)\)",
                RegexOptions.IgnoreCase);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells that contain formulas
                foreach (Cell cell in sheet.Cells)
                {
                    if (!cell.IsFormula)
                        continue;

                    string formula = cell.Formula;
                    Match match = textJoinRegex.Match(formula);
                    if (!match.Success)
                        continue; // No TEXTJOIN in this formula

                    // Extract delimiter (keep quotes as is)
                    string delimiter = match.Groups["delim"].Value.Trim();

                    // Extract source argument (could be a range or a list of arguments)
                    string source = match.Groups["source"].Value.Trim();

                    string newFormula;

                    // Simple handling for a single range like A1:A3
                    if (source.Contains(":"))
                    {
                        // Create a range object to enumerate cells
                        Aspose.Cells.Range range = sheet.Cells.CreateRange(source);
                        int cellCount = range.RowCount * range.ColumnCount;

                        StringBuilder sb = new StringBuilder();
                        sb.Append("="); // start formula
                        sb.Append("CONCATENATE(");

                        int added = 0;
                        for (int r = 0; r < range.RowCount; r++)
                        {
                            for (int c = 0; c < range.ColumnCount; c++)
                            {
                                // Reference to the cell (e.g., A1)
                                string cellRef = range[r, c].Name;

                                sb.Append(cellRef);
                                added++;

                                // Append delimiter between cells (except after the last one)
                                if (added < cellCount)
                                {
                                    sb.Append(","); // separator for CONCATENATE arguments
                                    sb.Append(delimiter);
                                    sb.Append(",");
                                }
                            }
                        }

                        sb.Append(")");
                        newFormula = sb.ToString();
                    }
                    else
                    {
                        // Source is a list of arguments (e.g., A1,B1,C1)
                        string[] args = source.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        StringBuilder sb = new StringBuilder();
                        sb.Append("=");
                        sb.Append("CONCATENATE(");
                        for (int i = 0; i < args.Length; i++)
                        {
                            sb.Append(args[i].Trim());
                            if (i < args.Length - 1)
                            {
                                sb.Append(","); // separator for CONCATENATE arguments
                                sb.Append(delimiter);
                                sb.Append(",");
                            }
                        }
                        sb.Append(")");
                        newFormula = sb.ToString();
                    }

                    // Replace the original formula with the new CONCATENATE formula
                    cell.Formula = newFormula;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
