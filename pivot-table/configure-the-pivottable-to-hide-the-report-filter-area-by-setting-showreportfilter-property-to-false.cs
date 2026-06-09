using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

namespace AsposeCellsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Get the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Hide the report filter area if the API supports it.
                // The ShowReportFilter property is not available in some versions,
                // so we use reflection to set it when present.
                try
                {
                    var showReportFilterProp = typeof(PivotTable).GetProperty("ShowReportFilter");
                    if (showReportFilterProp != null && showReportFilterProp.CanWrite)
                    {
                        showReportFilterProp.SetValue(pivotTable, false);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unable to modify report filter visibility: {ex.Message}");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}