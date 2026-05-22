using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

class Program
{
    static void Main()
    {
        // Load the workbook (load rule)
        string inputPath = "Input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Prepare the log file
        string logPath = "TrendlineLog.txt";
        File.WriteAllText(logPath, $"Trendline log generated at {DateTime.Now}{Environment.NewLine}");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts in the worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Process only column or bar charts (bar chart series)
                if (chart.Type == ChartType.Column || chart.Type == ChartType.Bar)
                {
                    // Add a linear trendline to each series
                    for (int i = 0; i < chart.NSeries.Count; i++)
                    {
                        // Add trendline (add rule)
                        int trendlineIndex = chart.NSeries[i].TrendLines.Add(TrendlineType.Linear);
                        Trendline trendline = chart.NSeries[i].TrendLines[trendlineIndex];

                        // Configure trendline to display equation and R‑squared
                        trendline.DisplayEquation = true;
                        trendline.DisplayRSquared = true;
                        trendline.Name = $"Series{i + 1} Linear Trend";

                        // Log the equation information (type and name)
                        string chartTitle = chart.Title?.Text ?? "Untitled";
                        string logEntry = $"Worksheet: {sheet.Name}, Chart: {chartTitle}, Series: {i + 1}, Trendline: {trendline.Name}, Type: {trendline.Type}{Environment.NewLine}";
                        File.AppendAllText(logPath, logEntry);
                    }
                }
            }
        }

        // Save the modified workbook (save rule)
        string outputPath = "Output.xlsx";
        workbook.Save(outputPath);
    }
}