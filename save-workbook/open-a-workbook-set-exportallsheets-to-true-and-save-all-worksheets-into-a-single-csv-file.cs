using System;
using Aspose.Cells;

namespace AsposeCellsExportAllSheetsToCsv
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (can be .xlsx, .xls, etc.)
            string sourcePath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(sourcePath);

            // Create CSV save options and enable exporting all worksheets
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
            csvOptions.ExportAllSheets = true;

            // Save all worksheets into a single CSV file
            string outputPath = "output_all_sheets.csv";
            workbook.Save(outputPath, csvOptions);

            Console.WriteLine($"Workbook saved to CSV with all sheets: {outputPath}");
        }
    }
}