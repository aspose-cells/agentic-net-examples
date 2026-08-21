// Title: Reformat Excel Dates to ISO 8601 with Aspose.Cells .NET Regex Replace
// Description: Demonstrates how to use Aspose.Cells for .NET to locate date strings in MM/dd/yyyy or MM-dd-yyyy format across an entire workbook and replace them with ISO 8601 (yyyy‑MM‑dd) using ReplaceOptions with RegexKey enabled.
// Keywords: Aspose.Cells | .NET | C# | regex date conversion | ISO 8601 Excel | search and replace workbook | ReplaceOptions | regular expression Excel | date format standardization | Excel automation
// Common Searches: Aspose.Cells replace date format with regex | C# convert MM/dd/yyyy to ISO 8601 in Excel | How to use ReplaceOptions RegexKey in Aspose.Cells | Search and replace dates in all worksheets .NET | Excel date reformatting programmatically
// Developer Intent: Automatically transform every textual date in a workbook to ISO 8601 format using a single regex replace operation.
// Use Cases: Standardize dates embedded in report text before distribution | Enable correct chronological sorting and filtering in generated Excel files | Apply a global text transformation without iterating each cell manually | Prepare data for systems that require ISO‑8601 timestamps
// AI Prompts: Write C# code with Aspose.Cells that finds dates like MM/dd/yyyy or MM-dd-yyyy in any cell and rewrites them as yyyy‑MM‑dd using ReplaceOptions.RegexKey. | Explain step‑by‑step how to configure ReplaceOptions for regex replacement and verify the changes in the saved workbook. | Provide a concise guide to apply a regex‑based date conversion to every worksheet in an existing Excel workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells for .NET to locate date strings in MM/dd/yyyy or MM-dd-yyyy format across an entire workbook and replace them with ISO 8601 (yyyy‑MM‑dd) using ReplaceOptions with RegexKey enabled.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample cells containing dates in different textual formats
        sheet.Cells["A1"].PutValue("Report date: 12/31/2022");
        sheet.Cells["A2"].PutValue("Event on 01-15-2023");
        sheet.Cells["A3"].PutValue("No date here");

        // Regex pattern to match dates like MM/dd/yyyy or MM-dd-yyyy
        string pattern = @"\b(\d{2})[/-](\d{2})[/-](\d{4})\b";
        // Replacement string to convert to ISO 8601 (yyyy-MM-dd)
        string replacement = "$3-$1-$2";

        // Set replace options to treat the pattern as a regular expression
        ReplaceOptions options = new ReplaceOptions
        {
            RegexKey = true
        };

        // Perform the replacement across the entire workbook
        workbook.Replace(pattern, replacement, options);

        // Save the modified workbook
        workbook.Save("ReformattedDates.xlsx");
    }
}
