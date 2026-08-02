// Title: C# – Find Regex Matches in the "LogEntries" Named Range with Aspose.Cells for .NET
// Description: Loads an Excel workbook, retrieves the named range "LogEntries", builds a CellArea that mirrors the range limits, and uses FindOptions (RegexKey, LookInType.Values, LookAtType.EntireContent) to locate every cell whose value matches a regular‑expression pattern (e.g., YYYY‑MM‑DD dates). Matches are printed to the console and the workbook is saved. Includes error handling for missing files and absent named ranges.
// Keywords: Aspose.Cells | C# | .NET | regex search | named range | FindOptions | CellArea | Excel pattern matching | LogEntries range | GitHub example | code sample
// Common Searches: Aspose.Cells regex search in named range C# | Find cells by pattern Aspose.Cells .NET | Search Excel named range with regular expression | Aspose.Cells FindOptions example GitHub | C# locate dates in a specific named range using Aspose
// Developer Intent: Identify all cells whose content matches a given regular expression inside the "LogEntries" named range of an Excel workbook.
// Use Cases: Extract every date string from a log worksheet for reporting or further analysis. | Validate that entries in a specific named range follow a required format before exporting data. | Highlight or log cells that meet a pattern criterion to support data quality audits.
// AI Prompts: Generate C# code that replaces regex matches within the "LogEntries" named range using Aspose.Cells and writes the updated values back to the workbook. | Show how to collect matched cells into a list, then create a new worksheet that lists each cell address and its matched value. | Provide an example of configuring FindOptions to locate email addresses in a named range and apply a background color to the matching cells.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Loads an Excel workbook, retrieves the named range "LogEntries", builds a CellArea that mirrors the range limits, and uses FindOptions (RegexKey, LookInType.Values, LookAtType.EntireContent) to locate every cell whose value matches a regular‑expression pattern (e.g., YYYY‑MM‑DD dates). Matches are printed to the console and the workbook is saved. Includes error handling for missing files and absent named ranges.
class Program
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "LogEntries"
            AsposeRange logRange = workbook.Worksheets.GetRangeByName("LogEntries");
            if (logRange == null)
            {
                Console.WriteLine("Named range 'LogEntries' not found.");
                return;
            }

            // Worksheet that contains the named range
            Worksheet sheet = logRange.Worksheet;

            // Configure find options for regex search within the named range
            FindOptions findOptions = new FindOptions
            {
                RegexKey = true,                     // Enable regex interpretation
                LookInType = LookInType.Values,      // Search cell values
                LookAtType = LookAtType.EntireContent // Exact match (no extra wildcards)
            };

            // Define the search area based on the named range
            CellArea area = new CellArea
            {
                StartRow = logRange.FirstRow,
                StartColumn = logRange.FirstColumn,
                EndRow = logRange.FirstRow + logRange.RowCount - 1,
                EndColumn = logRange.FirstColumn + logRange.ColumnCount - 1
            };
            findOptions.SetRange(area);

            // Regular expression pattern to look for (e.g., dates like 2023-07-28)
            string pattern = @"\d{4}-\d{2}-\d{2}";

            // Iterate through all matches in the named range
            Cell? previous = null;
            while (true)
            {
                Cell found = sheet.Cells.Find(pattern, previous, findOptions);
                if (found == null)
                    break;

                Console.WriteLine($"Found match at {found.Name}: {found.StringValue}");
                previous = found; // Continue searching after the current match
            }

            // Save the workbook (if any modifications were made)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
