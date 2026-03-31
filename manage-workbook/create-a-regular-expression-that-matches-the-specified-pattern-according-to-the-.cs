using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data in column A
        worksheet.Cells["A1"].PutValue("apple");
        worksheet.Cells["A2"].PutValue("banana");
        worksheet.Cells["A3"].PutValue("orange");
        worksheet.Cells["A4"].PutValue("pear");
        worksheet.Cells["A5"].PutValue("plum");

        // Configure replace options to treat the search key as a regular expression
        ReplaceOptions options = new ReplaceOptions
        {
            RegexKey = true,                 // Enable regex matching
            MatchEntireCellContents = false  // Allow partial matches within cell content
        };

        // Regular expression: match words that start with 'p' or 'b' followed by a vowel
        string regexPattern = "^[pb][aeiou]";

        // Replacement text
        string replacement = "FRUIT";

        // Perform the replacement using the regex pattern and options
        workbook.Replace(regexPattern, replacement, options);

        // Save the modified workbook
        workbook.Save("RegexReplaceOutput.xlsx");
    }
}