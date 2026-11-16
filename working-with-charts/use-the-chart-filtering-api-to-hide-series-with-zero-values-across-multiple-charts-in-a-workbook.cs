// Title: Programmatically hide chart series with all zero values in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide chart series that contain only zeros | filter out all-zero data series from multiple Excel charts using Aspose.Cells | set series.IsFiltered for zero-value series across worksheets in .NET | remove empty series from charts in an Excel file with Aspose.Cells API | iterate through workbook charts and hide zero value series C#
// Tags: chart series zero-value filtering Aspose.Cells | hide all-zero series .NET Excel | Aspose.Cells chart filtering API | iterate workbook charts C# | series.IsFiltered usage Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, iterates over every worksheet and each chart, checks each series' value range for all‑zero numeric data, marks such series as filtered using the chart filtering API (series.IsFiltered = true), and saves the modified workbook.
class HideZeroSeries
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts on the worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Iterate through each series in the chart
                for (int s = 0; s < chart.NSeries.Count; s++)
                {
                    Series series = chart.NSeries[s];

                    // Get the reference to the values range of the series (e.g., "Sheet1!A1:A5")
                    string valuesRef = series.Values;
                    if (string.IsNullOrEmpty(valuesRef))
                        continue; // No data source, skip

                    // Remove sheet name if present (e.g., "Sheet1!A1:A5" -> "A1:A5")
                    string cleanRef = valuesRef;
                    int exclPos = valuesRef.IndexOf('!');
                    if (exclPos >= 0 && exclPos < valuesRef.Length - 1)
                        cleanRef = valuesRef.Substring(exclPos + 1);

                    // Obtain the range object from the address string
                    Aspose.Cells.Range range;
                    try
                    {
                        range = sheet.Cells.CreateRange(cleanRef);
                    }
                    catch
                    {
                        // Invalid address, skip this series
                        continue;
                    }

                    // Determine if all numeric values in the range are zero
                    bool allZero = true;
                    foreach (Cell cell in range)
                    {
                        if (cell.Value == null)
                            continue;

                        if (double.TryParse(cell.Value.ToString(), out double d))
                        {
                            if (d != 0.0)
                            {
                                allZero = false;
                                break;
                            }
                        }
                        else
                        {
                            // Non‑numeric value, treat as non‑zero for safety
                            allZero = false;
                            break;
                        }
                    }

                    // Hide the series if all its values are zero
                    if (allZero)
                    {
                        series.IsFiltered = true; // Hide series using chart filtering API
                    }
                }
            }
        }

        try
        {
            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
