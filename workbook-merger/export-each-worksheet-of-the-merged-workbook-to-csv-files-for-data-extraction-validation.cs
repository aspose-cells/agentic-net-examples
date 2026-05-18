using System;
using Aspose.Cells;

namespace AsposeCellsExportEachSheetToCsv
{
    class Program
    {
        static void Main()
        {
            // Path to the merged workbook (result of previous merge operation)
            string mergedWorkbookPath = "MergedOutput.xlsx";

            // Load the merged workbook
            Workbook workbook = new Workbook(mergedWorkbookPath);

            // Prepare CSV save options – ExportAllSheets is false (default) so only the active sheet is saved
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
            csvOptions.ExportAllSheets = false; // ensure only the active sheet is exported

            // Export each worksheet to a separate CSV file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Set the current worksheet as active
                workbook.Worksheets.ActiveSheetIndex = i;

                // Build a file name for the CSV (e.g., Sheet1.csv, Sheet2.csv, ...)
                string csvFileName = $"Sheet{i + 1}.csv";

                // Save the active worksheet as CSV
                workbook.Save(csvFileName, csvOptions);
            }

            Console.WriteLine("All worksheets have been exported to individual CSV files.");
        }
    }
}