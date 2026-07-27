using System;
using Aspose.Cells;

class ReplaceWithRegexDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data containing placeholders like {{Name}} and {{Date}}
        sheet.Cells["A1"].PutValue("Hello {{Name}}");
        sheet.Cells["A2"].PutValue("Your appointment is on {{Date}}.");
        sheet.Cells["A3"].PutValue("{{Name}} has a meeting at {{Time}}.");

        // Define a regex pattern that matches placeholders of the form {{Word}}
        string pattern = @"\{\{[A-Za-z]+\}\}";
        string replacement = "REPLACED";

        // Set replace options to treat the search key as a regular expression
        ReplaceOptions options = new ReplaceOptions
        {
            RegexKey = true,               // Enable regex matching
            CaseSensitive = false,         // Case‑insensitive matching
            MatchEntireCellContents = false // Allow partial matches within cells
        };

        // Perform the replacement across the entire workbook
        int replacedCount = workbook.Replace(pattern, replacement, options);
        Console.WriteLine($"Total replacements made: {replacedCount}");

        // Save the workbook
        workbook.Save("RegexReplaceDemo.xlsx");
    }
}