// Title: Export Excel to CSV without Removing Blank Rows or Columns – Aspose.Cells C#
// Description: Loads an Excel workbook and saves it as CSV using Aspose.Cells with TxtSaveOptions (TrimLeadingBlankRowAndColumn = false, KeepSeparatorsForBlankRow = true) to preserve the original column layout and blank rows.
// Keywords: Aspose.Cells | CSV export | C# | TrimLeadingBlankRowAndColumn | KeepSeparatorsForBlankRow | preserve blank rows | preserve column structure | Excel to CSV | .NET | TxtSaveOptions
// Common Searches: Aspose.Cells export CSV keep empty rows | C# save workbook as CSV without trimming blanks | TxtSaveOptions TrimLeadingBlankRowAndColumn false example | How to retain column positions when converting Excel to CSV | KeepSeparatorsForBlankRow true usage
// Developer Intent: Convert an Excel workbook to a CSV file while maintaining all empty rows and columns so the output matches the source layout.
// Use Cases: Generating CSV reports from templates that contain placeholder rows or columns. | Creating data feeds where the row count must mirror the original worksheet, including fully blank rows. | Exporting spreadsheets with pre‑allocated blank columns for downstream systems that rely on fixed column positions.
// AI Prompts: Write C# code with Aspose.Cells to save a workbook as CSV using TrimLeadingBlankRowAndColumn = false and KeepSeparatorsForBlankRow = true. | Explain the impact of TrimLeadingBlankRowAndColumn and KeepSeparatorsForBlankRow on CSV output when converting Excel files. | Provide a step‑by‑step tutorial for converting Excel to CSV while preserving empty rows and columns using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // Loads an Excel workbook and saves it as CSV using Aspose.Cells with TxtSaveOptions (TrimLeadingBlankRowAndColumn = false, KeepSeparatorsForBlankRow = true) to preserve the original column layout and blank rows.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (any supported Excel format)
            string sourcePath = "input.xlsx";

            // Path for the resulting CSV file
            string csvPath = "output.csv";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Configure CSV save options:
            // - TrimLeadingBlankRowAndColumn = false ensures that leading empty rows/columns
            //   are preserved, keeping the original column structure unchanged.
            // - KeepSeparatorsForBlankRow = true retains separators for completely blank rows,
            //   so the row count in the CSV matches the worksheet.
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                TrimLeadingBlankRowAndColumn = false,
                KeepSeparatorsForBlankRow = true
            };

            // Save the workbook as CSV using the configured options
            workbook.Save(csvPath, csvOptions);

            Console.WriteLine($"Workbook successfully converted to CSV at: {csvPath}");
        }
    }
}
