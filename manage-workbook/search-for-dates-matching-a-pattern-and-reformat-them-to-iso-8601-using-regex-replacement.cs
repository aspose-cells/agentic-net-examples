using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data containing dates in MM/dd/yyyy format
        sheet.Cells["A1"].PutValue("01/15/2023");
        sheet.Cells["A2"].PutValue("12/31/2022");
        sheet.Cells["A3"].PutValue("No date here");
        sheet.Cells["A4"].PutValue("07/04/2021");

        // Regex pattern to match dates in MM/dd/yyyy
        string pattern = @"\b(\d{2})/(\d{2})/(\d{4})\b";

        // Replacement string to convert to ISO 8601 (yyyy-MM-dd) using captured groups
        string replacement = "${3}-${1}-${2}";

        // Configure replace options to treat the pattern as a regular expression
        ReplaceOptions options = new ReplaceOptions
        {
            RegexKey = true   // Enable regex matching
        };

        // Perform the replacement across the entire workbook
        workbook.Replace(pattern, replacement, options);

        // Save the modified workbook
        workbook.Save("ReformattedDates.xlsx");
    }
}