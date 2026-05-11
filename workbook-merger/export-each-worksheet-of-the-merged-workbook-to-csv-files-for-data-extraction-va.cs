using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace ExportWorksheetsToCsv
{
    class Program
    {
        static void Main()
        {
            // Paths of source Excel files to be merged
            string[] sourceFiles = { "File1.xlsx", "File2.xlsx" };
            // Temporary cache file required by CellsHelper.MergeFiles
            string cacheFile = "Cache.tmp";
            // Path for the merged workbook
            string mergedFile = "MergedWorkbook.xlsx";

            // Merge the source files into a single workbook
            CellsHelper.MergeFiles(sourceFiles, cacheFile, mergedFile);

            // Load the merged workbook
            Workbook workbook = new Workbook(mergedFile);

            // Directory where individual CSV files will be saved
            string outputDir = "CsvOutputs";
            Directory.CreateDirectory(outputDir);

            // Loop through each worksheet and export it to a separate CSV file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Set the current worksheet as active
                workbook.Worksheets.ActiveSheetIndex = i;

                // Configure CSV save options to export only the active sheet
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    ExportAllSheets = false   // Export only the active worksheet
                };

                // Build the CSV file name using the worksheet name
                string sheetName = workbook.Worksheets[i].Name;
                string csvPath = Path.Combine(outputDir, $"{sheetName}.csv");

                // Save the active worksheet as CSV
                workbook.Save(csvPath, csvOptions);
            }

            // Clean up temporary files (optional)
            if (File.Exists(cacheFile)) File.Delete(cacheFile);
            if (File.Exists(mergedFile)) File.Delete(mergedFile);
        }
    }
}