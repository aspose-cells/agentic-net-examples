// Title: How to delete chart series with incomplete data ranges using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates over each series in the first chart, examines the Values range for any blank or null cells, and removes the series that are not fully populated before saving the file. | Create a reusable C# method that takes a Chart object and eliminates all series whose source cells contain empty strings or null values, using Aspose.Cells range traversal. | Demonstrate a runtime check that evaluates the completeness of each chart series in Aspose.Cells and hides or deletes those series that have missing data, preserving the remaining series.
// Common Searches: asp.net remove chart series with blank cells using aspose.cells | c# check if chart series data range contains empty cells asp.net | how to delete incomplete series from an Excel chart with aspose.cells | runtime hide chart series based on missing data in aspose.cells .net | filter out chart series with null values in Excel using aspose.cells c#
// Tags: remove chart series Aspose.Cells C# | check series data completeness Aspose.Cells | delete series with empty cells Excel .NET | runtime chart series visibility Aspose.Cells | iterate chart series range Aspose.Cells | filter incomplete chart series C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, accesses the first worksheet and its first chart, iterates through each series to evaluate the cells referenced by the series' Values range, flags series containing any null or blank cells as incomplete, removes those series from the chart, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook contains no worksheets.");
                return;
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            Chart chart = worksheet.Charts[0];

            // Keep track of series indices that should be removed (incomplete data)
            List<int> indicesToRemove = new List<int>();

            // Iterate through each series in the chart
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                Series series = chart.NSeries[i];
                bool isComplete = true;

                // Get the range that defines the series values (e.g., "A1:B5")
                string valuesAddress = series.Values;

                if (string.IsNullOrWhiteSpace(valuesAddress))
                {
                    // If the series has no data source, consider it incomplete
                    isComplete = false;
                }
                else
                {
                    // Create a range object from the address
                    AsposeRange range = worksheet.Cells.CreateRange(valuesAddress);

                    // Check each cell in the range for emptiness
                    foreach (Cell cell in range)
                    {
                        // A cell is considered empty if its value is null or its string representation is empty
                        if (cell.Value == null || string.IsNullOrWhiteSpace(cell.StringValue))
                        {
                            isComplete = false;
                            break;
                        }
                    }
                }

                // If the series is incomplete, mark it for removal
                if (!isComplete)
                {
                    indicesToRemove.Add(i);
                }
            }

            // Remove incomplete series starting from the highest index to avoid shifting
            for (int i = indicesToRemove.Count - 1; i >= 0; i--)
            {
                chart.NSeries.RemoveAt(indicesToRemove[i]);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
