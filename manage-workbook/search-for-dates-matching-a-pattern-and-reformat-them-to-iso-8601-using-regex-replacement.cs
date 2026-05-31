using System;
using Aspose.Cells;

namespace AsposeCellsDateRegexReplace
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data containing dates in MM/dd/yyyy format
            cells["A1"].PutValue("Report date: 12/31/2022");
            cells["A2"].PutValue("Start: 01/15/2023 End: 02/20/2023");
            cells["A3"].PutValue("No date here");
            cells["A4"].PutValue("Another date 07/04/2021");

            // Define a regex pattern that matches dates in MM/dd/yyyy
            string datePattern = @"(\d{2})/(\d{2})/(\d{4})";

            // Replacement string to convert to ISO 8601 (yyyy-MM-dd)
            // $1 = month, $2 = day, $3 = year
            string isoReplacement = "$3-$1-$2";

            // Configure replace options to treat the pattern as a regular expression
            ReplaceOptions options = new ReplaceOptions
            {
                RegexKey = true,          // Enable regex matching
                CaseSensitive = false,    // Not relevant for dates but keep default
                MatchEntireCellContents = false // Replace dates within larger strings
            };

            // Perform the replacement across the entire workbook (lifecycle rule: replace)
            workbook.Replace(datePattern, isoReplacement, options);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("DateRegexReplaced.xlsx");
        }
    }
}