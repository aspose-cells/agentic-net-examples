// Title: Save Aspose.Cells Workbook as CSV while preserving blank rows (C#)
// Description: Demonstrates how to export a workbook to CSV using Aspose.Cells with TxtSaveOptions configured to keep separators for empty rows, disable trimming of leading blanks, and use UTF‑8 encoding with a comma delimiter.
// Keywords: Aspose.Cells CSV export | KeepSeparatorsForBlankRow | TxtSaveOptions | preserve empty rows CSV | TrimLeadingBlankRowAndColumn false | C# Aspose.Cells example | CSV delimiter configuration
// Common Searches: Aspose.Cells keep blank rows in CSV | How to export CSV with empty rows using Aspose | TxtSaveOptions KeepSeparatorsForBlankRow example | Prevent trimming leading blanks CSV Aspose.Cells | C# save workbook as CSV with separators
// Developer Intent: Export the entire workbook to a CSV file while retaining blank rows by inserting column separators for those rows.
// Use Cases: Generating CSV reports that include visual gaps for section separation. | Creating import files where blank rows act as logical delimiters between data blocks. | Preserving spreadsheet layout in CSV for downstream processing that relies on row positioning.
// AI Prompts: Write C# code with Aspose.Cells to save a workbook as CSV, preserving blank rows and using a custom delimiter. | Explain how KeepSeparatorsForBlankRow and TrimLeadingBlankRowAndColumn affect the CSV output in Aspose.Cells. | Adapt the example to use a semicolon delimiter and UTF‑16 encoding.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExample
{
    // Demonstrates how to export a workbook to CSV using Aspose.Cells with TxtSaveOptions configured to keep separators for empty rows, disable trimming of leading blanks, and use UTF‑8 encoding with a comma delimiter.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Add data with intentional blank rows
            cells[0, 0].PutValue("Header1");
            cells[0, 1].PutValue("Header2");
            cells[1, 0].PutValue("Row1Col1");
            cells[1, 1].PutValue("Row1Col2");
            // Row 2 is left blank
            cells[3, 0].PutValue("Row3Col1");
            cells[3, 1].PutValue("Row3Col2");

            // Configure CSV save options to keep separators for blank rows
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.UTF8,
                KeepSeparatorsForBlankRow = true, // Preserve empty rows with separators
                Separator = ',',                     // Standard CSV delimiter
                TrimLeadingBlankRowAndColumn = false // Ensure leading blanks are not trimmed
            };

            // Save the workbook as CSV using the configured options
            workbook.Save("output_with_blank_rows.csv", csvOptions);
        }
    }
}
