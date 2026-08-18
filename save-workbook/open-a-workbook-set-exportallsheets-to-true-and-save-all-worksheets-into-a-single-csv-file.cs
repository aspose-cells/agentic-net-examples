// Title: Aspose.Cells C# – Export All Workbook Sheets to a Single CSV File
// Description: Loads an Excel workbook with Aspose.Cells, configures TxtSaveOptions for CSV, enables ExportAllSheets, and saves every worksheet into one combined CSV file using .NET.
// Keywords: Aspose.Cells | C# | ExportAllSheets | TxtSaveOptions | CSV export | multiple worksheets to CSV | save workbook as CSV | Aspose.Cells .NET example | combine Excel sheets CSV | Aspose.Cells CSV all sheets
// Common Searches: Aspose.Cells export all sheets to CSV | C# save Excel workbook as one CSV file | TxtSaveOptions ExportAllSheets true example | combine Excel worksheets into a single CSV using Aspose | Aspose.Cells CSV export multiple sheets .NET
// Developer Intent: Export every worksheet in a workbook into one combined CSV file.
// Use Cases: Consolidate data from several sheets for analytics pipelines. | Generate a single CSV report from a multi‑sheet financial workbook. | Prepare a unified CSV for data migration when all sheets must be merged.
// AI Prompts: Write C# code with Aspose.Cells that loads an .xlsx file and saves all its worksheets into one CSV, ensuring ExportAllSheets is enabled. | Show how to configure TxtSaveOptions for CSV export with ExportAllSheets true and handle dynamic input/output paths. | Explain the behavior of ExportAllSheets in Aspose.Cells and describe the structure of the resulting CSV when multiple sheets are combined.

using System;
using Aspose.Cells;

namespace AsposeCellsExportAllSheetsToCsv
{
    // Loads an Excel workbook with Aspose.Cells, configures TxtSaveOptions for CSV, enables ExportAllSheets, and saves every worksheet into one combined CSV file using .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (can be .xlsx, .xls, etc.)
            string sourcePath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(sourcePath);

            // Create CSV save options and enable exporting all worksheets
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
            saveOptions.ExportAllSheets = true;

            // Save all worksheets into a single CSV file
            string outputPath = "output_all_sheets.csv";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to CSV with all sheets: {outputPath}");
        }
    }
}
