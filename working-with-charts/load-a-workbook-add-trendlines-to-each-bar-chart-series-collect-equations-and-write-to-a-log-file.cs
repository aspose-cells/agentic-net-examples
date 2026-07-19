// Title: Add Linear Trendlines to Bar & Column Charts and Log Details with Aspose.Cells for C#/.NET
// Description: Loads an Excel workbook, scans every worksheet for bar or column charts, adds a linear trendline to each series, shows the equation and R‑squared value, writes chart‑series‑trendline information to a log file, and saves the updated file.
// Keywords: Aspose.Cells C# trendline | add linear trendline bar chart | column chart trendline Aspose | log trendline equation .NET | chart series trendline automation | Excel trendline equation logging | Aspose.Cells chart manipulation | C# write log file
// Common Searches: how to add linear trendlines to all bar charts using Aspose.Cells | Aspose.Cells C# log trendline equation and R-squared | programmatically add trendlines to column charts in .NET | save workbook after modifying chart trendlines Aspose.Cells
// Developer Intent: Programmatically insert linear trendlines into every series of bar or column charts in an Excel workbook and record each trendline’s properties (type, name, equation, R‑squared) to a log file.
// Use Cases: Create a financial report that includes trendline equations for sales data visualized as column charts. | Automate quality‑control dashboards by adding trendlines to production bar charts and logging R‑squared values for audit trails. | Prepare data for downstream statistical analysis by enriching workbooks with trendlines while maintaining a change‑log for compliance.
// AI Prompts: Generate C# code using Aspose.Cells to add exponential trendlines to all line charts and export each equation to a CSV file. | Modify the provided code to set the trendline line color to red and increase its thickness for better visibility. | Explain how to retrieve the calculated equation string from a Trendline object after adding it to a chart with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an Excel workbook, scans every worksheet for bar or column charts, adds a linear trendline to each series, shows the equation and R‑squared value, writes chart‑series‑trendline information to a log file, and saves the updated file.
class TrendlineProcessor
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Prepare a log file (overwrite if exists)
        string logPath = "trendlines.log";
        File.WriteAllText(logPath, $"Trendline processing started at {DateTime.Now}{Environment.NewLine}");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts in the worksheet
            for (int chartIdx = 0; chartIdx < sheet.Charts.Count; chartIdx++)
            {
                Chart chart = sheet.Charts[chartIdx];

                // Process only Bar/Column charts
                if (chart.Type == ChartType.Column || chart.Type == ChartType.Bar)
                {
                    // Iterate through each series in the chart
                    for (int seriesIdx = 0; seriesIdx < chart.NSeries.Count; seriesIdx++)
                    {
                        // Add a linear trendline to the current series
                        int trendlineIndex = chart.NSeries[seriesIdx].TrendLines.Add(TrendlineType.Linear);
                        Trendline trendline = chart.NSeries[seriesIdx].TrendLines[trendlineIndex];

                        // Configure trendline to show equation and R‑squared value
                        trendline.DisplayEquation = true;
                        trendline.DisplayRSquared = true;
                        trendline.Name = $"Series{seriesIdx}_Linear";

                        // Log information about the added trendline
                        string logEntry = $"Worksheet: {sheet.Name}, Chart: {chartIdx}, Series: {seriesIdx}, " +
                                          $"Trendline Type: {trendline.Type}, Name: {trendline.Name}, " +
                                          $"DisplayEquation: {trendline.DisplayEquation}, DisplayRSquared: {trendline.DisplayRSquared}";
                        File.AppendAllText(logPath, logEntry + Environment.NewLine);
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");

        // Final log entry
        File.AppendAllText(logPath, $"Trendline processing completed at {DateTime.Now}{Environment.NewLine}");
    }
}
