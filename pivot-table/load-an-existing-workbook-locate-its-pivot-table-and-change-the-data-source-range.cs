using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing workbook that contains a pivot table
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(inputPath);

            // Assume the pivot table is in the first worksheet; adjust index if needed
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table (or locate by name if required)
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Define the new data source range.
            // The array contains the source range and the sheet name.
            // Example: data in C1:D10 on the same sheet ("Sheet1")
            string[] newDataSource = new string[] { "C1:D10", worksheet.Name };

            // Change the data source of the pivot table (uses the provided method)
            pivotTable.ChangeDataSource(newDataSource);

            // Refresh and recalculate the pivot table to apply the new source
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook (uses the provided save rule)
            string outputPath = "OutputWorkbook.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Pivot table data source changed and workbook saved to '{outputPath}'.");
        }
    }
}