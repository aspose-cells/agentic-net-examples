using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "PivotSource.xlsx";
                const string destPath = "PivotConsolidated.xlsx";

                // Ensure the source file exists
                if (!File.Exists(sourcePath))
                    throw new FileNotFoundException($"Source file not found: {sourcePath}");

                // Load the workbook
                var loadOptions = new LoadOptions();
                var workbook = new Workbook(sourcePath, loadOptions);

                // Access the first worksheet and its first pivot table
                var worksheet = workbook.Worksheets[0];
                if (worksheet.PivotTables.Count == 0)
                    throw new InvalidOperationException("No pivot tables found in the first worksheet.");

                var pivotTable = worksheet.PivotTables[0];

                // Consolidate report filter pages onto a single worksheet if the property exists
                var prop = typeof(PivotTable).GetProperty("ShowReportFilterPages");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(pivotTable, false);
                }

                // Save the modified workbook
                workbook.Save(destPath);
                Console.WriteLine($"Workbook saved to {destPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}