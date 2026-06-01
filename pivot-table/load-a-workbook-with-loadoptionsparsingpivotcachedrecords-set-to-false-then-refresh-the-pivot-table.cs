using System;
using Aspose.Cells;

namespace AsposeCellsPivotRefreshExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook that contains pivot tables
            string inputPath = "input.xlsx";

            // Create LoadOptions and disable parsing of pivot cached records
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.ParsingPivotCachedRecords = false;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Refresh all pivot tables in the workbook
            workbook.Worksheets.RefreshPivotTables();

            // Save the updated workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}