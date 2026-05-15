using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data containing dates in different formats
        sheet.Cells["A1"].PutValue("12/31/2022"); // MM/dd/yyyy
        sheet.Cells["A2"].PutValue("15-01-2023"); // dd-MM-yyyy
        sheet.Cells["A3"].PutValue("No date here");

        // Regex pattern to match dates like MM/dd/yyyy or dd-MM-yyyy
        string pattern = @"(\d{2})[\/-](\d{2})[\/-](\d{4})";

        // Replacement string to convert to ISO 8601 (yyyy-MM-dd)
        string replacement = "$3-$1-$2";

        // Configure replace options to treat the pattern as a regular expression
        ReplaceOptions options = new ReplaceOptions
        {
            RegexKey = true
        };

        // Perform the regex replacement across the entire workbook
        workbook.Replace(pattern, replacement, options);

        // Save the workbook with reformatted dates
        workbook.Save("ReformattedDates.xlsx");
    }
}