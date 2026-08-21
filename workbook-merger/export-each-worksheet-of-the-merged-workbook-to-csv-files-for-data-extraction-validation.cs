// Title: C# – Export Each Worksheet of a Merged Workbook to Separate CSV Files with Aspose.Cells
// Description: Merge multiple Excel files using CellsHelper.MergeFiles, load the combined workbook, iterate its worksheets, and save each one as an individual CSV file with TxtSaveOptions (ExportAllSheets = false). Optionally delete the temporary cache file.
// Keywords: Aspose.Cells CSV export | merge Excel files .NET | export worksheet to CSV C# | TxtSaveOptions ExportAllSheets | split merged workbook into CSV | Aspose.Cells workbook to CSV files
// Common Searches: Aspose.Cells export each sheet to CSV | C# merge Excel files and create CSV per worksheet | How to save a single worksheet as CSV using Aspose.Cells | Remove temporary cache file after CellsHelper.MergeFiles | Export active worksheet to CSV with TxtSaveOptions
// Developer Intent: Generate a separate CSV file for every worksheet in a merged Excel workbook.
// Use Cases: Validate data extraction by converting each merged sheet to CSV for downstream processing. | Produce per‑sheet CSV reports after consolidating client spreadsheets. | Automate archival of individual worksheet snapshots for audit compliance.
// AI Prompts: Write C# code that merges a list of Excel files with Aspose.Cells and exports each worksheet to its own CSV file, ensuring only the active sheet is saved. | Show how to configure TxtSaveOptions in Aspose.Cells to export a single worksheet as CSV (ExportAllSheets = false). | Explain how to clean up the temporary cache file created by CellsHelper.MergeFiles before exporting worksheets.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Merge multiple Excel files using CellsHelper.MergeFiles, load the combined workbook, iterate its worksheets, and save each one as an individual CSV file with TxtSaveOptions (ExportAllSheets = false). Optionally delete the temporary cache file.
class ExportWorksheetsToCsv
{
    static void Main()
    {
        // Paths of source Excel files to be merged
        string[] sourceFiles = new string[] { "File1.xlsx", "File2.xlsx" };

        // Temporary cache file required by CellsHelper.MergeFiles
        string cacheFile = "CacheFile.tmp";

        // Destination merged workbook file
        string mergedFile = "MergedWorkbook.xlsx";

        // Merge the source files into a single workbook
        CellsHelper.MergeFiles(sourceFiles, cacheFile, mergedFile);

        // Load the merged workbook
        Workbook mergedWorkbook = new Workbook(mergedFile);

        // Export each worksheet to a separate CSV file
        for (int i = 0; i < mergedWorkbook.Worksheets.Count; i++)
        {
            // Set the current worksheet as active
            mergedWorkbook.Worksheets.ActiveSheetIndex = i;

            // Build a CSV file name that includes the worksheet name
            string sheetName = mergedWorkbook.Worksheets[i].Name;
            string csvFileName = Path.GetFileNameWithoutExtension(mergedFile) + "_" + sheetName + ".csv";

            // Configure TxtSaveOptions to export only the active sheet
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                ExportAllSheets = false   // ensure only the active sheet is saved
            };

            // Save the active worksheet as CSV
            mergedWorkbook.Save(csvFileName, saveOptions);
        }

        // Clean up temporary files (optional)
        if (File.Exists(cacheFile))
        {
            File.Delete(cacheFile);
        }
    }
}
