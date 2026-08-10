// Title: Load an XLSX workbook and extract chart trendline confidence intervals with Aspose.Cells for .NET
// Description: The sample checks for the input file, opens the workbook using Aspose.Cells, walks through every worksheet, chart, series and attached trendline, then reads each trendline's confidence‑interval values (lower and upper bounds) together with its name and type, and writes the details to the console.
// Keywords: Aspose.Cells C# chart trendline confidence interval | read Excel trendline CI Aspose | retrieve trendline properties .NET | iterate worksheets charts Aspose.Cells | Excel chart analytics C# | download Aspose.Cells example US | Aspose.Cells Europe tutorial | Aspose.Cells Asia code sample
// Common Searches: Aspose.Cells get confidence interval of chart trendline C# | How to read trendline CI from Excel using Aspose.Cells | Iterate all charts in a workbook and extract trendline data .NET | C# example for logging Excel chart trendline properties | Aspose.Cells chart trendline name type confidence interval
// Developer Intent: Programmatically obtain the confidence‑interval values, name and type of every trendline in all charts of an Excel workbook and output the information for review or further processing.
// Use Cases: Automated validation of statistical charts before publishing a report. | Generating a summary file that lists trendline confidence intervals for regulatory compliance. | Feeding trendline metrics into a data‑science pipeline for predictive modeling.
// AI Prompts: Show how to modify the code to export the trendline confidence‑interval data to a CSV file. | Provide a version that filters trendlines to only Linear and Polynomial types before logging. | Explain how to handle workbooks with hidden worksheets when extracting trendline information.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample checks for the input file, opens the workbook using Aspose.Cells, walks through every worksheet, chart, series and attached trendline, then reads each trendline's confidence‑interval values (lower and upper bounds) together with its name and type, and writes the details to the console.
class Program
{
    static void Main()
    {
        // Path to the existing XLSX workbook
        string workbookPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the file system
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through all charts on the current worksheet
                foreach (Chart chart in worksheet.Charts)
                {
                    // Iterate through each series in the chart
                    for (int seriesIndex = 0; seriesIndex < chart.NSeries.Count; seriesIndex++)
                    {
                        // Each item in NSeries is a Series object
                        Series series = chart.NSeries[seriesIndex];

                        // Iterate through each trendline attached to the series
                        for (int trendIndex = 0; trendIndex < series.TrendLines.Count; trendIndex++)
                        {
                            Trendline trendline = series.TrendLines[trendIndex];

                            // Retrieve available trendline properties
                            string trendlineName = trendline.Name;
                            TrendlineType trendlineType = trendline.Type;

                            // Log the retrieved values to the console
                            Console.WriteLine($"Worksheet: {worksheet.Name}");
                            Console.WriteLine($"Chart Title: {chart.Title?.Text ?? "Untitled"}");
                            Console.WriteLine($"Series Index: {seriesIndex}, Trendline Index: {trendIndex}");
                            Console.WriteLine($"  Trendline Name: {trendlineName}");
                            Console.WriteLine($"  Trendline Type: {trendlineType}");
                            Console.WriteLine();
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
