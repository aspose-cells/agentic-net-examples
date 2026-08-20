// Title: Replace placeholders with regex across an entire Aspose.Cells workbook (C#)
// Description: Shows how to create a workbook, insert cells with {{placeholder}} tokens, enable RegexKey in ReplaceOptions, define a pattern (\{\{[A-Za-z0-9_]+\}\}), run Workbook.Replace to swap all matches with a custom value, and save the result.
// Keywords: Aspose.Cells regex replace | C# replace placeholders workbook | ReplaceOptions RegexKey | bulk text replace Excel Aspose | replace all worksheets Aspose.Cells | Excel placeholder substitution .NET
// Common Searches: Aspose.Cells replace regex across workbook | C# replace {{placeholder}} in Excel using Aspose | How to use ReplaceOptions RegexKey in Aspose.Cells | Bulk replace text in all sheets Aspose.Cells .NET
// Developer Intent: Replace every placeholder that matches a regex pattern throughout the entire workbook.
// Use Cases: Clean template files by swapping {{...}} tokens with a marker before data population. | Mass‑update dynamic fields such as order IDs or dates in generated reports with a single call. | Mask confidential placeholders across all worksheets in one operation.
// AI Prompts: Write C# code that uses Aspose.Cells to perform a case‑sensitive regex replacement on all cells while preserving formatting. | Explain how to configure ReplaceOptions to replace only whole‑cell matches of a pattern across multiple worksheets. | Provide an example that iterates through worksheets and applies different regex replacements for separate placeholder groups using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsReplaceRegexDemo
{
    // Shows how to create a workbook, insert cells with {{placeholder}} tokens, enable RegexKey in ReplaceOptions, define a pattern (\{\{[A-Za-z0-9_]+\}\}), run Workbook.Replace to swap all matches with a custom value, and save the result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data containing placeholders
            sheet.Cells["A1"].PutValue("Hello {{Name}}");
            sheet.Cells["A2"].PutValue("Your order {{OrderId}} is confirmed.");
            sheet.Cells["A3"].PutValue("Contact {{Email}} for support.");

            // Configure replace options to treat the search key as a regular expression
            ReplaceOptions options = new ReplaceOptions
            {
                RegexKey = true,          // Enable regex matching
                CaseSensitive = false,    // Optional: ignore case
                MatchEntireCellContents = false // Allow partial matches within cells
            };

            // Define a regex pattern that matches placeholders like {{Placeholder}}
            string placeholderPattern = @"\{\{[A-Za-z0-9_]+\}\}";

            // Perform the replacement across the entire workbook (lifecycle rule: replace)
            workbook.Replace(placeholderPattern, "REPLACED", options);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ReplacedPlaceholders.xlsx");
        }
    }
}
