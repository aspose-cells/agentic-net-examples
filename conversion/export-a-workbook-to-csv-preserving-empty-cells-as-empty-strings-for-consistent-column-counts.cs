// Title: Aspose.Cells C# – Export Workbook to CSV While Keeping Empty Cells Intact
// Description: Demonstrates how to save an Aspose.Cells workbook as a UTF‑8 CSV using TxtSaveOptions. The settings KeepSeparatorsForBlankRow, TrimLeadingBlankRowAndColumn, and TrimTrailingBlankCells ensure that blank cells are written as empty strings, preserving column alignment across all rows.
// Keywords: Aspose.Cells CSV export C# | keep empty cells CSV Aspose | TxtSaveOptions KeepSeparatorsForBlankRow | TrimLeadingBlankRowAndColumn Aspose | TrimTrailingBlankCells CSV | UTF-8 CSV Aspose.Cells | fixed column count Excel to CSV | Aspose.Cells conversion example
// Common Searches: Aspose.Cells export to CSV without removing empty cells | How to keep column count when saving Excel as CSV in C# | CSV options to preserve blank cells Aspose | Save workbook as UTF‑8 CSV using Aspose.Cells
// Developer Intent: Generate a CSV file from a workbook where every cell, including blanks, is represented so that each row has the same number of columns.
// Use Cases: Creating CSV reports that must retain placeholder fields for downstream processing. | Exporting price lists where some items lack values but column structure cannot change. | Producing data feeds for systems that expect a fixed column count per record.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to CSV, ensuring empty cells are output as empty strings. | Explain the impact of KeepSeparatorsForBlankRow, TrimLeadingBlankRowAndColumn, and TrimTrailingBlankCells on CSV output. | Suggest additional TxtSaveOptions for custom delimiters, quoting, or line endings when exporting to CSV.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // Demonstrates how to save an Aspose.Cells workbook as a UTF‑8 CSV using TxtSaveOptions. The settings KeepSeparatorsForBlankRow, TrimLeadingBlankRowAndColumn, and TrimTrailingBlankCells ensure that blank cells are written as empty strings, preserving column alignment across all rows.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some data with intentional empty cells
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Apple");
            // B2 is left empty on purpose
            cells["B3"].PutValue(2.99);

            // Create CSV save options
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                // Ensure separators are written for completely blank rows
                KeepSeparatorsForBlankRow = true,
                // Do not trim leading blank rows/columns so empty cells stay as empty strings
                TrimLeadingBlankRowAndColumn = false,
                // Do not trim trailing blank cells in a row
                TrimTailingBlankCells = false,
                // Use UTF-8 encoding for the CSV file
                Encoding = Encoding.UTF8
            };

            // Save the workbook as CSV with the specified options
            workbook.Save("output.csv", csvOptions);
        }
    }
}
