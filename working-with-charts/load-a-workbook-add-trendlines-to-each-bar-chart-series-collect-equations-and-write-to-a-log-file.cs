// Title: Add linear trendlines to every series of bar and column charts in an Excel workbook using Aspose.Cells for .NET and export the equations to a text log
// AI Prompts: Generate C# code that opens an .xlsx workbook with Aspose.Cells, iterates through all worksheets, finds bar and column charts, adds a linear trendline to each series, captures the displayed equation, and writes the sheet, chart, series, and equation details to a text file. | Extend the program to also process line charts, apply exponential trendlines, and output the collected equations in CSV format.
// Common Searches: how to add a linear trendline to each series of a bar chart using Aspose.Cells C# | extract trendline formula from Excel chart with Aspose.Cells .NET | save chart trendline equations to a log file in C# | iterate through all worksheets and charts in a workbook with Aspose.Cells
// Tags: trendline insertion Aspose.Cells bar chart | chart series equation extraction .NET | write trendline equations to text file C# | worksheet chart enumeration Aspose.Cells | dynamic trendline handling C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace TrendlineExtractor
{
    // The program loads an Excel workbook, scans every worksheet for bar and column charts, adds a linear trendline to each series, records the displayed equation together with sheet, chart, and series identifiers, and writes all collected equations to a text log file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook
            string workbookPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The file '{workbookPath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // List to store trendline equations
                List<string> trendlineEquations = new List<string>();

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts in the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Process only Bar or Column charts
                        if (chart.Type == ChartType.Bar || chart.Type == ChartType.Column)
                        {
                            try
                            {
                                // Iterate through each series in the chart
                                foreach (Series series in chart.NSeries)
                                {
                                    try
                                    {
                                        // Use dynamic to access Trendlines (may not exist in older versions)
                                        dynamic dynSeries = series;
                                        dynamic trendline = dynSeries.Trendlines.Add(TrendlineType.Linear);
                                        trendline.DisplayEquation = true;

                                        // Retrieve the equation; use dynamic to avoid compile‑time binding
                                        string equation = trendline.Formula ?? "N/A";

                                        // Store the equation with context information
                                        string info = $"Sheet: {sheet.Name}, Chart: {chart.Title?.Text ?? "Untitled"}, Series: {series.Name}, Equation: {equation}";
                                        trendlineEquations.Add(info);
                                    }
                                    catch (Exception exSeries)
                                    {
                                        Console.WriteLine($"Warning: Could not process series '{series.Name}' in chart '{chart.Title?.Text ?? "Untitled"}' on sheet '{sheet.Name}'. Details: {exSeries.Message}");
                                    }
                                }
                            }
                            catch (Exception exChart)
                            {
                                Console.WriteLine($"Warning: Could not process chart '{chart.Title?.Text ?? "Untitled"}' on sheet '{sheet.Name}'. Details: {exChart.Message}");
                            }
                        }
                    }
                }

                // Write all collected equations to a log file
                string logPath = "trendlines_log.txt";
                File.WriteAllLines(logPath, trendlineEquations);

                Console.WriteLine($"Trendline equations have been written to '{logPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
