using System;
using Aspose.Cells;

namespace AsposeCellsLoadFilterDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (replace with an actual file path)
            string sourceFile = "sample.xlsx";

            // Create a LoadFilter that loads only cell values and formulas (no formatting, charts, etc.)
            LoadFilter filter = new LoadFilter(LoadDataFilterOptions.CellValue | LoadDataFilterOptions.Formula);

            // Assign the filter to LoadOptions
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = filter;

            // Load the workbook using the file path and the configured LoadOptions
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Display basic information about the loaded workbook
            Console.WriteLine("Number of worksheets loaded: " + workbook.Worksheets.Count);
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet '{sheet.Name}' contains {sheet.Cells.Count} cells (filtered).");
            }

            // Save the filtered workbook to a new file
            workbook.Save("FilteredOutput.xlsx", SaveFormat.Xlsx);
        }
    }
}