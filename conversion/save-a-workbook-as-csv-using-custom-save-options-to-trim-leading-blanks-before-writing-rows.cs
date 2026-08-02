// Title: Aspose.Cells C# – Save Workbook to CSV while removing leading blank rows and columns
// Description: Creates a workbook, places data starting at C3, configures TxtSaveOptions to discard initial empty rows/columns, sets a comma separator and UTF‑8 encoding, and saves the result as trimmed_output.csv.
// Keywords: Aspose.Cells | C# CSV export | TrimLeadingBlankRowAndColumn | TxtSaveOptions | remove leading blanks | custom CSV delimiter | UTF-8 CSV | sparse worksheet export | save workbook as CSV | Aspose.Cells example
// Common Searches: Aspose.Cells trim leading blanks CSV | TxtSaveOptions TrimLeadingBlankRowAndColumn C# | export worksheet to CSV without empty rows Aspose | set CSV delimiter and encoding Aspose.Cells | remove initial empty rows columns when saving CSV
// Developer Intent: Export a worksheet to CSV while automatically eliminating the empty rows and columns that appear before the first populated cell.
// Use Cases: Generate compact CSV files from sheets where data begins beyond A1, avoiding leading commas. | Produce UTF‑8 encoded CSVs with a specific delimiter for downstream systems that cannot handle blank cells. | Simplify processing of sparse spreadsheets by stripping pre‑data empty rows/columns during export.
// AI Prompts: Show how to also trim trailing empty rows and columns when saving a workbook as CSV with Aspose.Cells. | Provide code to export every worksheet in a workbook to separate CSV files using the same trimming and encoding settings. | Explain the impact of the TrimLeadingBlankRowAndColumn property on CSV output and any scenarios where it does not apply.

using System;
using System.Text;
using Aspose.Cells;

// Creates a workbook, places data starting at C3, configures TxtSaveOptions to discard initial empty rows/columns, sets a comma separator and UTF‑8 encoding, and saves the result as trimmed_output.csv.
class CsvTrimLeadingBlanksExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add data with leading blank rows and columns (e.g., start at C3)
        cells["C3"].PutValue("Data1");
        cells["D4"].PutValue("Data2");
        cells["E5"].PutValue("Data3");

        // Create TxtSaveOptions and enable trimming of leading blank rows/columns
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.TrimLeadingBlankRowAndColumn = true; // Trim leading blanks like Excel
        saveOptions.Separator = ',';                     // Use comma as CSV delimiter
        saveOptions.Encoding = Encoding.UTF8;            // Set desired encoding

        // Save the workbook as CSV using the custom options
        workbook.Save("trimmed_output.csv", saveOptions);
    }
}
