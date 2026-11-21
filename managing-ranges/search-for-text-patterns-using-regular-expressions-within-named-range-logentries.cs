// Title: Find cells that start with "ERROR:" using a regular expression in the "LogEntries" named range with Aspose.Cells for .NET (C#)
// Common Searches: how to use Regex to locate error messages in a named range with Aspose.Cells C# | search for cells beginning with 'ERROR:' in an Excel named range using Aspose.Cells | iterate through Aspose.Cells named range and extract matching strings in C# | retrieve row and column indices of cells that match a pattern in an Aspose.Cells workbook
// Tags: regex matching in Aspose.Cells named range | GetRangeByName cell iteration C# | extract error messages from Excel cells using Aspose.Cells | process string cells with compiled Regex Aspose.Cells | log entries pattern search Aspose.Cells workbook

using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook, obtains the "LogEntries" named range, and uses a compiled regular expression to identify cells whose text starts with "ERROR:". Matching cells are recorded with their row and column positions, printed to the console, and the workbook is saved.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the range defined by the named range "LogEntries"
            // GetRangeByName returns an Aspose.Cells.Range object
            Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName("LogEntries");
            if (range == null)
            {
                Console.WriteLine("Named range 'LogEntries' not found.");
                return;
            }

            // Define the regex pattern to find lines starting with "ERROR:"
            string pattern = @"\bERROR:\s*(.+)";
            Regex regex = new Regex(pattern, RegexOptions.Compiled);

            // List to hold matches (row, column, matched text)
            List<(int Row, int Column, string Match)> matches = new List<(int, int, string)>();

            // Iterate through each cell in the named range
            foreach (Cell cell in range)
            {
                // Process only string cells
                if (cell.Type == CellValueType.IsString)
                {
                    string text = cell.StringValue;
                    foreach (Match m in regex.Matches(text))
                    {
                        matches.Add((cell.Row, cell.Column, m.Value));
                    }
                }
            }

            // Output the matches
            foreach (var item in matches)
            {
                Console.WriteLine($"Cell[{item.Row},{item.Column}] : {item.Match}");
            }

            // Save the workbook (even if unchanged)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
