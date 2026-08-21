// Title: C# – Replace CONCATENATE with CONCAT in an Excel workbook using Aspose.Cells
// Description: Load an existing workbook, configure a case‑insensitive ReplaceOptions object, substitute every legacy CONCATENATE function with the modern CONCAT function, and save the updated file. The example demonstrates bulk formula migration across all worksheets with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# Excel formula replace | CONCATENATE to CONCAT | ReplaceOptions | case insensitive replace | .NET Excel automation | bulk formula update | Excel workbook manipulation | GitHub Aspose.Cells example | Aspose.Cells API
// Common Searches: How to replace CONCATENATE with CONCAT using Aspose.Cells C# | Aspose.Cells case‑insensitive formula replacement | Bulk update Excel formulas across a workbook .NET | Replace legacy Excel functions with Aspose.Cells | Convert CONCATENATE to CONCAT programmatically
// Developer Intent: Swap every occurrence of the CONCATENATE function for CONCAT throughout an Excel workbook via Aspose.Cells.
// Use Cases: Modernize legacy spreadsheets before distribution by converting old CONCATENATE formulas to the newer CONCAT syntax. | Automate large‑scale report generation where all worksheets must use the current CONCAT function. | Apply a case‑insensitive text replacement to rename deprecated Excel functions in password‑protected or multi‑sheet workbooks. | Create a reusable utility for bulk formula migration in CI/CD pipelines that handle Excel assets.
// AI Prompts: Write C# code with Aspose.Cells that replaces a list of deprecated Excel functions (e.g., CONCATENATE, TEXTJOIN) with their modern equivalents, allowing the caller to set case sensitivity and target specific worksheets. | Show how to iterate through each worksheet, detect formula cells containing a given function name, replace it with a new function, and save the workbook while preserving protection settings.

using System;
using Aspose.Cells;

// Load an existing workbook, configure a case‑insensitive ReplaceOptions object, substitute every legacy CONCATENATE function with the modern CONCAT function, and save the updated file. The example demonstrates bulk formula migration across all worksheets with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace "input.xlsx" with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure replace options to ignore case (so CONCATENATE, Concatenate, etc. are all matched)
        ReplaceOptions replaceOptions = new ReplaceOptions
        {
            CaseSensitive = false,
            MatchEntireCellContents = false
        };

        // Replace every occurrence of the old function name with the new one
        workbook.Replace("CONCATENATE", "CONCAT", replaceOptions);

        // Save the updated workbook (replace "output.xlsx" with desired output path)
        workbook.Save("output.xlsx");
    }
}
