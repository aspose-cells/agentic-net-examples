using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace ExportWorksheetsToCsv
{
    class Program
    {
        static void Main()
        {
            // Path to the merged workbook (result of previous merge operation)
            string mergedWorkbookPath = "MergedWorkbook.xlsx";

            // Load the merged workbook
            Workbook workbook = new Workbook(mergedWorkbookPath);

            // Iterate through all worksheets in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Set the current worksheet as the active sheet
                workbook.Worksheets.ActiveSheetIndex = i;

                // Create CSV save options; ExportAllSheets = false (default) ensures only the active sheet is saved
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
                saveOptions.ExportAllSheets = false;

                // Build a file name for the current sheet (e.g., Sheet1.csv, Sheet2.csv, ...)
                string csvFileName = $"Sheet{i + 1}.csv";

                // Save the active worksheet to a CSV file
                workbook.Save(csvFileName, saveOptions);
            }

            Console.WriteLine("All worksheets have been exported to individual CSV files.");
        }
    }
}