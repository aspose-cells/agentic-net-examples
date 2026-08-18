// Title: Aspose.Cells C# – Save Workbook to CSV while Trimming Leading Blank Rows/Columns
// Description: Shows how to build a workbook, place data at C3, and use TxtSaveOptions (comma delimiter, UTF‑8 encoding, TrimLeadingBlankRowAndColumn = true) to export a CSV that excludes any initial empty rows or columns.
// Keywords: Aspose.Cells | C# | CSV export | TxtSaveOptions | TrimLeadingBlankRowAndColumn | remove leading blanks | UTF-8 CSV | custom delimiter | Excel to CSV .NET | save workbook as CSV
// Common Searches: Aspose.Cells trim leading blanks CSV | TxtSaveOptions TrimLeadingBlankRowAndColumn example | export Excel to CSV without empty rows C# | remove initial empty columns when saving CSV Aspose | custom CSV delimiter Aspose.Cells
// Developer Intent: Export an Excel workbook to CSV while automatically discarding leading empty rows and columns.
// Use Cases: Produce clean CSV reports from spreadsheets that contain header rows after blank rows or columns. | Feed CSV files into data pipelines that cannot handle leading empty rows. | Generate UTF‑8 encoded CSV files with a specific delimiter and no preceding blanks.
// AI Prompts: Write C# code with Aspose.Cells to save a workbook as a semicolon‑delimited CSV and trim leading empty rows and columns. | Explain the effect of the TrimLeadingBlankRowAndColumn property in TxtSaveOptions on the output CSV. | Show how to configure TxtSaveOptions for UTF‑16 encoding and keep leading blanks unchanged.

using System;
using Aspose.Cells;
using System.Text;

namespace AsposeCellsCsvTrimExample
{
    // Shows how to build a workbook, place data at C3, and use TxtSaveOptions (comma delimiter, UTF‑8 encoding, TrimLeadingBlankRowAndColumn = true) to export a CSV that excludes any initial empty rows or columns.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data with leading blank rows and columns
            // Row 0 and column 0 are left blank intentionally
            cells["C3"].PutValue("Data1"); // Row index 2, Column index 2
            cells["D4"].PutValue("Data2"); // Row index 3, Column index 3
            cells["E5"].PutValue("Data3"); // Row index 4, Column index 4

            // Configure text (CSV) save options
            TxtSaveOptions saveOptions = new TxtSaveOptions();
            saveOptions.Separator = ',';                     // Use comma as delimiter
            saveOptions.Encoding = Encoding.UTF8;            // UTF-8 encoding
            saveOptions.TrimLeadingBlankRowAndColumn = true; // Trim leading blank rows/columns

            // Save the workbook as CSV using the custom options
            workbook.Save("TrimmedOutput.csv", saveOptions);

            Console.WriteLine("Workbook saved as CSV with leading blanks trimmed.");
        }
    }
}
