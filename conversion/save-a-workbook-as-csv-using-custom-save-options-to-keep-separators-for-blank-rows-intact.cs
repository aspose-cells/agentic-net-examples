// Title: Save Workbook as CSV with Blank‑Row Separators Using Aspose.Cells TxtSaveOptions (C#)
// Description: Creates a workbook, inserts data with an empty row, and saves it as a CSV file using TxtSaveOptions. The options set a comma separator, enable KeepSeparatorsForBlankRow, and use UTF‑8 encoding, ensuring that blank rows retain column delimiters.
// Keywords: Aspose.Cells CSV export | TxtSaveOptions KeepSeparatorsForBlankRow | custom CSV separator C# | UTF-8 CSV Aspose.Cells | preserve blank rows CSV
// Common Searches: Aspose.Cells keep commas for empty rows CSV | TxtSaveOptions KeepSeparatorsForBlankRow example C# | save Excel as CSV with custom delimiter Aspose | export workbook to CSV preserving blank lines
// Developer Intent: Export an Excel workbook to CSV while keeping delimiter characters in rows that contain no data.
// Use Cases: Generating fixed‑width CSV files required by legacy import tools. | Creating data feeds where each row must contain the same number of columns, even when some rows are intentionally empty. | Producing CSV reports that preserve row alignment for downstream processing.
// AI Prompts: Show how to change TxtSaveOptions to use a semicolon as the CSV separator while keeping blank‑row delimiters. | Provide a C# snippet that reads a CSV saved with KeepSeparatorsForBlankRow and validates the number of commas per row. | Explain the impact of KeepSeparatorsForBlankRow on CSV output when the source workbook contains merged cells.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExample
{
    // Creates a workbook, inserts data with an empty row, and saves it as a CSV file using TxtSaveOptions. The options set a comma separator, enable KeepSeparatorsForBlankRow, and use UTF‑8 encoding, ensuring that blank rows retain column delimiters.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Populate data with intentional blank rows
            cells[0, 0].PutValue("First");
            cells[0, 1].PutValue("Row");
            // Row 1 is left blank
            cells[2, 0].PutValue("Third");
            cells[2, 1].PutValue("Row");

            // Configure CSV save options
            TxtSaveOptions csvOptions = new TxtSaveOptions
            {
                // Use comma as separator
                Separator = ',',
                // Ensure blank rows retain separators
                KeepSeparatorsForBlankRow = true,
                // Use UTF-8 encoding for the output file
                Encoding = Encoding.UTF8
            };

            // Save the workbook as CSV with the custom options
            workbook.Save("output.csv", csvOptions);

            Console.WriteLine("Workbook saved as CSV with separators for blank rows.");
        }
    }
}
