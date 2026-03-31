using System;
using Aspose.Cells;

namespace AsposeCellsRegexReplace
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data that will be processed by the regex replacement
            sheet.Cells["A1"].PutValue("Order #12345");
            sheet.Cells["A2"].PutValue("Invoice #98765");
            sheet.Cells["A3"].PutValue("Reference #ABC123");
            sheet.Cells["B1"].PutValue("No numbers here");
            sheet.Cells["B2"].PutValue("Mixed 12AB34");

            // Define the regular expression pattern.
            // Example: replace any sequence of digits with the string "XXXX"
            string pattern = @"\d+";

            // Define the replacement string
            string replacement = "XXXX";

            // Configure replace options to treat the pattern as a regular expression
            ReplaceOptions options = new ReplaceOptions
            {
                RegexKey = true,          // Enable regex processing
                CaseSensitive = false,    // Case does not matter for this pattern
                MatchEntireCellContents = false // Allow partial matches within cell text
            };

            // Perform the regex replacement across all cells in the workbook
            // (lifecycle rule: replace)
            workbook.Replace(pattern, replacement, options);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("RegexReplacedOutput.xlsx");
        }
    }
}