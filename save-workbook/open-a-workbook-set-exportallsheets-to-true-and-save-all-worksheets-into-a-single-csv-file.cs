using System;
using Aspose.Cells;

namespace AsposeCellsExportAllSheetsToCsv
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Create CSV save options and enable exporting all worksheets
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
            csvOptions.ExportAllSheets = true;

            // Save all worksheets into a single CSV file
            workbook.Save("output_all_sheets.csv", csvOptions);
        }
    }
}