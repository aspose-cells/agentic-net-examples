// Title: Hide chart series that contain only zero values across all worksheets using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to walk through every chart in a workbook and set Series.IsFiltered = true when the series' data range consists exclusively of zeros. | Demonstrate how to apply the Aspose.Cells chart filtering API to automatically conceal zero‑value series in multiple worksheets.
// Common Searches: Aspose.Cells hide chart series with all zeros in C# | how to filter out zero‑value series in Excel charts using Aspose.Cells | iterate over all charts in a workbook and hide empty series Aspose.Cells | set Series.IsFiltered programmatically for zero data series in .NET
// Tags: chart series zero-value filtering Aspose.Cells | use Aspose.Cells API to hide series | iterate workbook charts C# | hide empty data series Excel | evaluate chart series data range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, loops through each worksheet and its charts, checks every series' value range, marks the series as filtered when all cells are zero or empty, and saves the updated workbook.
class HideZeroSeries
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts on the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Process each series in the chart
                    for (int i = 0; i < chart.NSeries.Count; i++)
                    {
                        Series series = chart.NSeries[i];

                        // Get the range string that contains the series values (e.g., "Sheet1!$B$2:$B$5")
                        string valuesRange = series.Values;
                        if (string.IsNullOrEmpty(valuesRange))
                            continue; // No values defined for this series

                        // Create a Range object from the values string
                        AsposeRange range = sheet.Cells.CreateRange(valuesRange);

                        bool allZero = true;

                        // Examine each cell in the range
                        foreach (Cell cell in range)
                        {
                            object valObj = cell.Value;
                            if (valObj == null)
                                continue; // Treat empty cells as zero

                            // Try to parse the cell value as a double
                            if (double.TryParse(valObj.ToString(), out double d))
                            {
                                if (d != 0.0)
                                {
                                    allZero = false;
                                    break; // Found a non‑zero value
                                }
                            }
                            else
                            {
                                // Non‑numeric value – treat as non‑zero
                                allZero = false;
                                break;
                            }
                        }

                        // Hide the series if all its values are zero
                        if (allZero)
                        {
                            series.IsFiltered = true; // Hides the series in the chart
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
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
}
