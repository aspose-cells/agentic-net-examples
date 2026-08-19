// Title: Case‑insensitive find and replace of "Revenue Total" in an Excel workbook using Aspose.Cells for .NET
// Description: The sample loads an Excel workbook with Aspose.Cells, configures ReplaceOptions to ignore case and allow partial matches, substitutes every occurrence of revenue‑related text (e.g., "total revenue", "Total Revenue") with the standardized phrase "Revenue Total", and saves the updated file.
// Keywords: Aspose.Cells case insensitive replace | C# Excel find and replace | Workbook.Replace options | standardize revenue terminology | Excel text normalization .NET | replace string in .xlsx | financial spreadsheet automation
// Common Searches: Aspose.Cells replace text ignoring case | C# replace partial cell content in Excel | How to standardize revenue labels in Excel using Aspose | Case‑insensitive string replace in .NET Excel library | Update legacy financial spreadsheets programmatically
// Developer Intent: Replace all variations of revenue‑related wording with the consistent label "Revenue Total" in an Excel file using Aspose.Cells for .NET.
// Use Cases: Ensure uniform financial terminology across quarterly reports before distribution. | Modernize legacy workbooks where revenue headings appear in mixed capitalizations. | Automate post‑processing of generated Excel files to enforce a single revenue label.
// AI Prompts: Generate C# code that replaces multiple revenue‑related phrases with "Revenue Total" across all worksheets using Aspose.Cells. | Explain how ReplaceOptions properties affect case‑insensitive matching in Workbook.Replace. | Show how to log each replacement made when normalizing text in an Excel workbook with Aspose.Cells.

using System;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, configures ReplaceOptions to ignore case and allow partial matches, substitutes every occurrence of revenue‑related text (e.g., "total revenue", "Total Revenue") with the standardized phrase "Revenue Total", and saves the updated file.
class ReplaceTotalRevenue
{
    static void Main()
    {
        // Load the workbook (replace "input.xlsx" with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure replace options for case‑insensitive search
        ReplaceOptions options = new ReplaceOptions
        {
            CaseSensitive = false,          // ignore case
            MatchEntireCellContents = false // allow partial matches within a cell
        };

        // Replace all occurrences of "total revenue" with "Revenue Total"
        workbook.Replace("total revenue", "Revenue Total", options);

        // Save the modified workbook (replace "output.xlsx" with your desired path)
        workbook.Save("output.xlsx");
    }
}
