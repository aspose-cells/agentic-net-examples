// Title: Hide or remove Excel chart series with incomplete data ranges using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates through each series in a chart, checks if the series' value range contains any empty cells, and sets the series IsVisible property to false when blanks are found. | Create a helper method in C# using Aspose.Cells that validates that all cells in a given range are populated, and demonstrate how to call it to conditionally set series visibility at runtime. | Show how to save the workbook after setting visibility for incomplete series and log which series were modified.
// Common Searches: Aspose.Cells C# filter out chart series with empty cells | programmatically check Excel chart series data completeness before rendering with Aspose.Cells | set IsVisible property of chart series based on range validation in Aspose.Cells .NET
// Tags: hide chart series Aspose.Cells C# | validate range completeness Aspose.Cells | conditional series visibility Excel .NET | remove incomplete series Aspose.Cells | runtime chart data check Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, examines each series in the first chart, uses a helper to verify that the series' value range has no empty cells, removes or hides any series that are incomplete, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = worksheet.Charts[0];

            // Iterate through each series in reverse order to allow removal
            for (int i = chart.NSeries.Count - 1; i >= 0; i--)
            {
                Series series = chart.NSeries[i];

                // Retrieve the address of the series values (e.g., "A1:A10")
                string valuesRange = series.Values;

                // Determine if the data range is complete (no empty cells)
                bool isComplete = IsRangeComplete(workbook, valuesRange);

                // If the range is incomplete, remove (hide) the series
                if (!isComplete)
                {
                    chart.NSeries.RemoveAt(i);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to verify that all cells in a given range contain data
    static bool IsRangeComplete(Workbook workbook, string rangeAddress)
    {
        try
        {
            // Create a Range object from the address on the first worksheet
            Aspose.Cells.Range range = workbook.Worksheets[0].Cells.CreateRange(rangeAddress);

            // Check each cell in the range
            foreach (Cell cell in range)
            {
                // If a cell is null or its string representation is empty, the range is incomplete
                if (cell.Value == null || string.IsNullOrWhiteSpace(cell.StringValue))
                    return false;
            }

            // All cells contain data
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error checking range '{rangeAddress}': {ex.Message}");
            return false;
        }
    }
}
