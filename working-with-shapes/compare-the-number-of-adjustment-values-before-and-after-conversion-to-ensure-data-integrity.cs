// Title: Check that chart series count stays the same after saving an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to open an existing .xlsx file, read the NSeries count of the first chart, save the workbook to a new file, reload it, and compare the two counts. | Create a reusable C# method that takes a workbook path, saves the workbook to another location with Aspose.Cells, and returns a boolean indicating whether the chart series count before and after the save are identical.
// Common Searches: aspnet compare chart series number before and after workbook save Aspose.Cells | c# verify that chart NSeries count does not change when re-saving an Excel file | how to ensure chart data integrity after converting XLSX with Aspose.Cells | Aspose.Cells check chart adjustments consistency after file export in C#
// Tags: Aspose.Cells chart series verification | detect series mismatch after workbook export | C# chart data consistency check | validate NSeries integrity on save | Excel chart preservation with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads an original XLSX workbook, records the NSeries count of its first chart, saves the workbook to a new file, reloads the saved file, and compares the series counts to confirm they remain unchanged.
class AdjustmentComparison
{
    static void Main()
    {
        // Paths for the original workbook and the converted workbook
        string originalPath = "original.xlsx";
        string convertedPath = "converted.xlsx";

        try
        {
            // Verify that the original file exists to avoid FileNotFoundException
            if (!File.Exists(originalPath))
            {
                Console.WriteLine($"File not found: {originalPath}");
                return;
            }

            // Load the original workbook
            Workbook originalWorkbook = new Workbook(originalPath);
            Worksheet originalSheet = originalWorkbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (originalSheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the original worksheet.");
                return;
            }

            // Get the first chart in the worksheet
            Chart originalChart = originalSheet.Charts[0];

            // Use NSeries count as a proxy for adjustment count
            int originalSeriesCount = originalChart.NSeries.Count;

            // Save the workbook to another format (placeholder conversion)
            originalWorkbook.Save(convertedPath, SaveFormat.Xlsx);

            // Verify that the converted file was created
            if (!File.Exists(convertedPath))
            {
                Console.WriteLine($"Failed to create converted file: {convertedPath}");
                return;
            }

            // Load the converted workbook
            Workbook convertedWorkbook = new Workbook(convertedPath);
            Worksheet convertedSheet = convertedWorkbook.Worksheets[0];

            // Ensure the converted worksheet contains at least one chart
            if (convertedSheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the converted worksheet.");
                return;
            }

            Chart convertedChart = convertedSheet.Charts[0];
            int convertedSeriesCount = convertedChart.NSeries.Count;

            // Compare the counts and output the result
            if (originalSeriesCount == convertedSeriesCount)
            {
                Console.WriteLine($"Series count matches: {originalSeriesCount}");
            }
            else
            {
                Console.WriteLine($"Series count mismatch. Before: {originalSeriesCount}, After: {convertedSeriesCount}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
