// Title: C# – Save Workbook as CSV with Blank‑Row Delimiters using Aspose.Cells TxtSaveOptions
// Description: Demonstrates how to create a workbook, insert data with intentional empty rows, and export it to CSV while preserving column separators for those rows. The example configures TxtSaveOptions with a custom separator, enables KeepSeparatorsForBlankRow, and sets the desired text encoding.
// Keywords: Aspose.Cells CSV export | TxtSaveOptions KeepSeparatorsForBlankRow | custom CSV separator .NET | blank row delimiters | CSV encoding Aspose.Cells | C# workbook to CSV
// Common Searches: Aspose.Cells keep commas for empty rows CSV | TxtSaveOptions KeepSeparatorsForBlankRow C# example | save workbook as CSV with custom separator Aspose.Cells | export CSV with ASCII encoding using Aspose.Cells | preserve blank rows in CSV export Aspose
// Developer Intent: Export a workbook to CSV while ensuring that completely empty rows still contain the defined column separators.
// Use Cases: Generate CSV reports that must retain placeholder delimiters for blank rows, keeping row count consistent for downstream parsers. | Create CSV files with specific encodings (e.g., ASCII, UTF‑8) for legacy systems without losing empty‑row structure. | Customize the CSV delimiter (comma, semicolon, tab) and preserve delimiters for empty rows when converting Excel data with Aspose.Cells.
// AI Prompts: Write C# code with Aspose.Cells to save a workbook as a semicolon‑separated CSV and keep delimiters for blank rows. | Show how to set TxtSaveOptions to UTF‑8 encoding and enable KeepSeparatorsForBlankRow for CSV export in Aspose.Cells. | Explain the purpose of KeepSeparatorsForBlankRow in Aspose.Cells CSV conversion and how it affects empty rows.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExample
{
    // Demonstrates how to create a workbook, insert data with intentional empty rows, and export it to CSV while preserving column separators for those rows. The example configures TxtSaveOptions with a custom separator, enables KeepSeparatorsForBlankRow, and sets the desired text encoding.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Populate some data with blank rows in between
            cells[0, 0].PutValue("First");
            cells[0, 1].PutValue("Row");
            // Row 1 is left blank intentionally
            cells[2, 0].PutValue("Third");
            cells[2, 1].PutValue("Row");

            // Configure TxtSaveOptions for CSV output
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                // Use comma as the CSV separator
                Separator = ',',
                // Ensure separators are written for completely blank rows
                KeepSeparatorsForBlankRow = true,
                // Use ASCII encoding for demonstration (any encoding can be used)
                Encoding = Encoding.ASCII
            };

            // Save the workbook as CSV using the configured options
            // This uses the Workbook.Save(string, SaveOptions) rule
            string outputPath = "output.csv";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to '{outputPath}' with separators kept for blank rows.");
        }
    }
}
