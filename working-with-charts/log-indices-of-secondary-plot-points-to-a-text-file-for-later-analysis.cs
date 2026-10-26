// Title: Logging secondary axis series indices from Excel charts to a text file with Aspose.Cells for .NET
// Common Searches: c# find series plotted on secondary axis using Aspose.Cells and export to txt | iterate Excel workbook charts and identify secondary axis series in Aspose.Cells | write secondary plot point details to a file with Aspose.Cells library | detect secondary axis series in a workbook and log to a text file in C# | Aspose.Cells example for secondary axis series detection and logging
// Tags: Aspose.Cells secondary axis series detection | export chart series indices to text file C# | iterate worksheets and charts with Aspose.Cells | reflection fallback for IsOnSecondaryAxis property | log secondary plot points in C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryPlotLogger
{
    // The sample loads an Excel workbook using Aspose.Cells, walks through each worksheet, chart, and series, determines whether a series is plotted on the secondary axis (checking both IsOnSecondaryAxis and PlotOnSecondAxis via reflection for version compatibility), and writes entries containing the worksheet name, chart name, and series index to a specified text file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be processed
            string excelPath = "InputWorkbook.xlsx";

            // Path to the output text file where indices will be logged
            string logPath = "SecondaryPlotPoints.txt";

            // Verify that the input workbook exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file '{excelPath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // List to hold logged entries
                List<string> logEntries = new List<string>();

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts in the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Iterate through all series in the chart
                        for (int i = 0; i < chart.NSeries.Count; i++)
                        {
                            Series series = chart.NSeries[i];

                            // Determine if the series is plotted on the secondary axis.
                            // Use reflection to stay compatible with different Aspose.Cells versions.
                            bool isOnSecondary = false;
                            var prop = series.GetType().GetProperty("IsOnSecondaryAxis");
                            if (prop != null && prop.PropertyType == typeof(bool))
                            {
                                isOnSecondary = (bool)prop.GetValue(series);
                            }
                            else
                            {
                                // Fallback to older property name if present
                                var altProp = series.GetType().GetProperty("PlotOnSecondAxis");
                                if (altProp != null && altProp.PropertyType == typeof(bool))
                                {
                                    isOnSecondary = (bool)altProp.GetValue(series);
                                }
                            }

                            if (isOnSecondary)
                            {
                                // Log worksheet name, chart name, and series index
                                string entry = $"Worksheet: {sheet.Name}, Chart: {chart.Name}, SeriesIndex: {i}";
                                logEntries.Add(entry);
                            }
                        }
                    }
                }

                // Write all logged entries to the text file
                File.WriteAllLines(logPath, logEntries);

                Console.WriteLine($"Logging complete. {logEntries.Count} secondary plot point(s) recorded to '{logPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
