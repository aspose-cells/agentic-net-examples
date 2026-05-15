using Aspose.Cells;
using System;

class ReplaceRegexDemo
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Regular expression pattern to match the placeholder (example: PLACEHOLDER_ followed by digits)
        string pattern = @"\bPLACEHOLDER_\d+\b";

        // Replacement text
        string replacement = "REPLACED";

        // Set up replace options to treat the pattern as a regular expression
        ReplaceOptions options = new ReplaceOptions
        {
            RegexKey = true,               // Enable regex matching
            CaseSensitive = false,         // Case‑insensitive match
            MatchEntireCellContents = false // Allow partial matches within cell contents
        };

        // Perform the replacement across the entire workbook
        int replacedCount = workbook.Replace(pattern, replacement, options);
        Console.WriteLine($"Replacements made: {replacedCount}");

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}