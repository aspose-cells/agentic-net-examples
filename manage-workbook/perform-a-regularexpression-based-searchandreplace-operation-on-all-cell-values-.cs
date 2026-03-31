using System;
using Aspose.Cells;

namespace AsposeCellsRegexReplace
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Define the regular expression pattern and the replacement text
            string regexPattern = @"\bApple\b";          // example: match the word "Apple"
            string replacementText = "Orange";

            // Configure replace options to treat the pattern as a regular expression
            ReplaceOptions options = new ReplaceOptions
            {
                RegexKey = true,          // enable regex processing
                CaseSensitive = false,    // ignore case (set true if needed)
                MatchEntireCellContents = false // allow partial matches within cell text
            };

            // Perform the regex replace across all worksheets and cells
            int replacedCount = workbook.Replace(regexPattern, replacementText, options);

            // Optionally, output the number of replacements made
            Console.WriteLine($"Total replacements performed: {replacedCount}");

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}