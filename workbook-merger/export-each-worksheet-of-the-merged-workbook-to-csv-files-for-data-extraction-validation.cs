// Title: Export worksheets from a merged workbook to separate CSV files using Aspose.Cells for .NET
// Description: Load a merged Excel workbook with Aspose.Cells, iterate through its worksheets, set each sheet as active, configure TxtSaveOptions to export only the active sheet, and save each one as an individual CSV file (e.g., Sheet_1.csv, Sheet_2.csv).
// Keywords: Aspose.Cells | C# | export worksheet to CSV | TxtSaveOptions | ExportAllSheets false | merged workbook | multiple CSV files | loop worksheets | save as CSV | Excel to CSV conversion
// Common Searches: Aspose.Cells export each sheet to CSV | C# save individual worksheets as CSV files | How to export only active worksheet with TxtSaveOptions | Convert merged Excel workbook to separate CSV files | Loop through workbook sheets and generate CSV in .NET
// Developer Intent: Create separate CSV files for every worksheet in a merged Excel workbook.
// Use Cases: Validate data extraction by comparing each sheet’s CSV output. | Provide per‑sheet CSV feeds for downstream analytics pipelines. | Generate reporting files after consolidating multiple workbooks.
// AI Prompts: Write C# code that loads a merged workbook with Aspose.Cells and exports each worksheet to a uniquely named CSV file. | Show how to set TxtSaveOptions.ExportAllSheets to false to save only the active sheet as CSV. | Explain how to modify the export loop to use the worksheet’s name instead of an index in the CSV filename.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace ExportWorksheetsToCsv
{
    // Load a merged Excel workbook with Aspose.Cells, iterate through its worksheets, set each sheet as active, configure TxtSaveOptions to export only the active sheet, and save each one as an individual CSV file (e.g., Sheet_1.csv, Sheet_2.csv).
    class Program
    {
        static void Main()
        {
            // Path to the merged workbook (created earlier by CellsHelper.MergeFiles or Workbook.Combine)
            string mergedWorkbookPath = "MergedOutput.xlsx";

            // Load the merged workbook
            Workbook workbook = new Workbook(mergedWorkbookPath);

            // Iterate through each worksheet and export it as an individual CSV file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Set the current worksheet as the active sheet
                workbook.Worksheets.ActiveSheetIndex = i;

                // Define the CSV file name for the current sheet
                string csvFileName = $"Sheet_{i + 1}.csv";

                // Configure TxtSaveOptions to export only the active sheet
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    ExportAllSheets = false   // Export only the active worksheet
                };

                // Save the active worksheet to CSV using the Save(string, SaveOptions) overload
                workbook.Save(csvFileName, saveOptions);
            }

            Console.WriteLine("All worksheets have been exported to separate CSV files.");
        }
    }
}
