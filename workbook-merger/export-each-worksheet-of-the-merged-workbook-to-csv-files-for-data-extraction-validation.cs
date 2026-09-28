// Title: Export each worksheet from a merged Excel workbook to separate CSV files with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a merged .xlsx workbook using Aspose.Cells, iterates through all worksheets, and saves each one as an individual .csv file named after the worksheet. | Show how to configure TxtSaveOptions to export only the active worksheet to CSV and how to change the active sheet inside a loop with Aspose.Cells. | Provide a C# example that sanitizes worksheet names for safe file creation and sets a custom CSV delimiter when exporting sheets from a merged workbook.
// Common Searches: Aspose.Cells .NET export each sheet of a combined workbook to separate CSV files | C# save individual worksheets as CSV from a merged Excel file using Aspose.Cells | How to loop through worksheets and export to CSV with TxtSaveOptions in Aspose.Cells | Export merged workbook worksheets to CSV for data validation using Aspose.Cells | Create CSV files per worksheet from a merged .xlsx using Aspose.Cells for .NET
// Tags: export worksheet to csv using TxtSaveOptions | iterate workbook worksheets Aspose.Cells .NET | merged workbook csv extraction Aspose.Cells | save active sheet as csv Aspose.Cells | custom csv delimiter Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace ExportWorksheetsToCsv
{
    // Loads a merged workbook (MergedWorkbook.xlsx) with Aspose.Cells, loops through each worksheet, sets it as the active sheet, and saves it as a CSV file named after the worksheet using TxtSaveOptions.
    class Program
    {
        static void Main()
        {
            // Path to the merged workbook (created earlier by CellsHelper.MergeFiles or Workbook.Combine)
            string mergedWorkbookPath = "MergedWorkbook.xlsx";

            // Load the merged workbook
            Workbook workbook = new Workbook(mergedWorkbookPath);

            // Iterate through each worksheet and export it to an individual CSV file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Set the current worksheet as the active sheet
                workbook.Worksheets.ActiveSheetIndex = i;

                // Build a CSV file name based on the worksheet name (or index)
                string sheetName = workbook.Worksheets[i].Name;
                string csvFileName = $"{sheetName}.csv";

                // Create TxtSaveOptions for CSV format.
                // ExportAllSheets defaults to false, which means only the active sheet will be saved.
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);

                // Save the active worksheet to CSV using the save options
                workbook.Save(csvFileName, saveOptions);

                Console.WriteLine($"Worksheet '{sheetName}' exported to '{csvFileName}'.");
            }
        }
    }
}
