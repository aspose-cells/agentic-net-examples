// Title: How to add linear trendlines with equation labels to each bar or column chart series in an Excel workbook using Aspose.Cells for .NET and record the actions in a log file
// AI Prompts: Write C# code that opens a specified .xlsx file with Aspose.Cells, iterates through all worksheets, identifies bar, stacked bar, column, and stacked column charts, adds a linear trendline to every series, enables equation display, logs the worksheet, chart, and series details, and saves the modified workbook to a new file. | Enhance the existing program to also handle line charts by adding a second‑order polynomial trendline, display its equation on the chart, and append appropriate entries to the same log file.
// Common Searches: aspnet add linear trendline to bar chart series using Aspose.Cells | c# log each trendline addition in an Excel workbook with Aspose.Cells | display trendline equation on bar charts via Aspose.Cells .NET | use reflection to add trendlines in older Aspose.Cells versions | iterate through all charts in a workbook and modify series Aspose.Cells
// Tags: linear trendline addition for bar and column charts Aspose.Cells | show trendline equation on Excel chart C# | log chart modifications Aspose.Cells workbook | reflection based trendline support Aspose.Cells | iterate worksheets and charts Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing Excel file with Aspose.Cells, creates a log file, and walks through each worksheet and its charts. For bar, stacked bar, column, and stacked column charts it adds a linear trendline to every series via reflection (to stay compatible with older library versions), turns on equation display, writes a detailed log entry, and finally saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";
            const string logFile = "trendlines_log.txt";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Prepare a log file to collect processing details
            using (StreamWriter logWriter = new StreamWriter(logFile))
            {
                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts in the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Process only Bar or Column type charts (including stacked variants)
                        if (chart.Type == ChartType.Bar ||
                            chart.Type == ChartType.BarStacked ||
                            chart.Type == ChartType.Column ||
                            chart.Type == ChartType.ColumnStacked)
                        {
                            // Iterate through each series in the chart
                            foreach (Series series in chart.NSeries)
                            {
                                try
                                {
                                    // Attempt to add a linear trendline if the API supports it
                                    // Note: Trendline support may vary by Aspose.Cells version.
                                    // The following block uses reflection to avoid compile‑time errors
                                    // on older versions that lack the Trendlines property.
                                    var trendlinesProp = series.GetType().GetProperty("Trendlines");
                                    if (trendlinesProp != null)
                                    {
                                        var trendlines = trendlinesProp.GetValue(series);
                                        var addMethod = trendlines?.GetType().GetMethod("Add");
                                        if (addMethod != null)
                                        {
                                            var trendlineObj = addMethod.Invoke(trendlines, new object[] { TrendlineType.Linear });
                                            // Enable equation display if possible
                                            var displayEqProp = trendlineObj?.GetType().GetProperty("DisplayEquation");
                                            displayEqProp?.SetValue(trendlineObj, true);
                                            logWriter.WriteLine(
                                                $"Worksheet: {sheet.Name}, Chart: {chart.Name}, Series: {series.Name}, Trendline: added");
                                        }
                                        else
                                        {
                                            logWriter.WriteLine(
                                                $"Worksheet: {sheet.Name}, Chart: {chart.Name}, Series: {series.Name}, Trendline: not supported (Add method missing)");
                                        }
                                    }
                                    else
                                    {
                                        logWriter.WriteLine(
                                            $"Worksheet: {sheet.Name}, Chart: {chart.Name}, Series: {series.Name}, Trendline: not supported (Trendlines property missing)");
                                    }
                                }
                                catch (Exception exSeries)
                                {
                                    Console.WriteLine($"Failed to process series '{series.Name}': {exSeries.Message}");
                                }
                            }
                        }
                    }
                }
            }

            // Save the workbook (unchanged if trendlines were not added)
            workbook.Save(outputFile);
            Console.WriteLine($"Processing completed. Output saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
