using System;
using Aspose.Cells;

namespace AsposeCellsReplaceRegexDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells with data that matches the regex pattern
            sheet.Cells["A1"].PutValue("Item-001");
            sheet.Cells["A2"].PutValue("Item-002");
            sheet.Cells["A3"].PutValue("Product-ABC");
            sheet.Cells["A4"].PutValue("Item-003");

            // Define the regular expression pattern to find all strings starting with "Item-"
            string regexPattern = @"Item-\d{3}";
            // Define the replacement text
            string replacement = "ReplacedItem";

            // Configure replace options to treat the search key as a regular expression
            ReplaceOptions options = new ReplaceOptions
            {
                RegexKey = true,          // Enable regex matching
                CaseSensitive = false,    // Case-insensitive matching
                MatchEntireCellContents = false // Allow partial matches within cell contents
            };

            // Perform the replacement across the entire workbook
            int replacedCount = workbook.Replace(regexPattern, replacement, options);
            Console.WriteLine($"Total replacements made: {replacedCount}");

            // Save the workbook to a file
            workbook.Save("RegexReplaceResult.xlsx");
        }
    }
}