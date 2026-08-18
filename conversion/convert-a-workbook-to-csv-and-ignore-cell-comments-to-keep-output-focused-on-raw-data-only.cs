// Title: Convert Excel to CSV without Comments using Aspose.Cells for .NET
// Description: Loads an .xlsx workbook, clears every worksheet's comments, strips personal information, configures TxtSaveOptions (ExportAllSheets, no separators for blank rows, trim leading blanks), and saves a single CSV file.
// Keywords: Aspose.Cells CSV conversion | remove Excel comments C# | clear worksheet comments | TxtSaveOptions CSV settings | export all sheets to CSV | trim blank rows CSV | remove personal information Aspose | C# Aspose.Cells export
// Common Searches: Aspose.Cells export to CSV without comments | C# remove cell comments before CSV conversion | How to clear Excel comments using Aspose.Cells | Save multiple worksheets as one CSV file .NET | Trim leading blank rows when saving CSV with Aspose
// Developer Intent: Generate a CSV file from an Excel workbook while discarding all cell comments and any embedded personal data.
// Use Cases: Produce a clean data feed for ETL pipelines from multi‑sheet workbooks. | Share spreadsheet content with external partners without exposing internal notes. | Create a single CSV snapshot for analytics tools, ensuring no empty separators appear.
// AI Prompts: Write C# code with Aspose.Cells that converts an .xlsx to CSV, removes all comments and personal info, and combines every sheet into one file. | Explain the effect of ExportAllSheets, KeepSeparatorsForBlankRow, and TrimLeadingBlankRowAndColumn on the CSV output after comments are cleared. | Adapt the example to export only the active worksheet while still ignoring comments.

using System;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // Loads an .xlsx workbook, clears every worksheet's comments, strips personal information, configures TxtSaveOptions (ExportAllSheets, no separators for blank rows, trim leading blanks), and saves a single CSV file.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Remove all comments from every worksheet to keep the CSV output focused on raw data
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.ClearComments();
            }

            // Optional: also remove any personal information that might be stored with comments
            workbook.RemovePersonalInformation();

            // Configure CSV (text) save options
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                // Export all sheets into a single CSV file (set to false to export only the active sheet)
                ExportAllSheets = true,

                // Ensure blank rows are not filled with separators (default is false)
                KeepSeparatorsForBlankRow = false,

                // Trim leading blank rows/columns as Excel does
                TrimLeadingBlankRowAndColumn = true
            };

            // Save the workbook as CSV
            string csvPath = "output.csv";
            workbook.Save(csvPath, csvOptions);

            Console.WriteLine($"Workbook has been converted to CSV without comments: {csvPath}");
        }
    }
}
