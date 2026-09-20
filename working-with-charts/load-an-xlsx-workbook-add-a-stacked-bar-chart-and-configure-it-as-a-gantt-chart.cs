// Title: Add a stacked bar chart to an existing XLSX workbook and format it as a Gantt chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing XLSX file with Aspose.Cells, inserts a horizontal stacked bar chart, and configures the series to represent start dates and durations for a Gantt chart. | Show how to bind task names from column A as category labels, hide the chart legend, and set a low gap width to make the bars appear thicker in an Aspose.Cells chart. | Demonstrate saving the modified workbook to a new file after the Gantt chart is created and include error handling for a missing input file.
// Common Searches: Aspose.Cells C# create Gantt chart from existing Excel workbook | how to add a stacked bar chart as a Gantt chart using Aspose.Cells .NET | set chart category axis labels from column A in Aspose.Cells | adjust gap width and hide legend in Aspose.Cells stacked bar chart | load workbook, insert chart, save new file Aspose.Cells example
// Tags: Aspose.Cells create Gantt chart | Aspose.Cells bind chart categories from worksheet column | Aspose.Cells hide chart legend | Aspose.Cells set chart gap width | C# load and modify XLSX workbook with Aspose.Cells | Aspose.Cells add bar chart to worksheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLSX file, adds a horizontal stacked bar chart with start and duration series to represent tasks, configures category labels, hides the legend, adjusts bar thickness, and saves the workbook as a new file.
class GanttChartExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or any target worksheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Add a stacked bar chart to the worksheet (horizontal stacked bar)
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title (optional)
            chart.Title.Text = "Gantt Chart";

            // Add a series for the start dates (used as offset)
            int startSeriesIndex = chart.NSeries.Add("B2:B5", true);
            Series startSeries = chart.NSeries[startSeriesIndex];
            startSeries.Name = "Start";

            // Add a series for the durations (the visible bars)
            int durationSeriesIndex = chart.NSeries.Add("C2:C5", true);
            Series durationSeries = chart.NSeries[durationSeriesIndex];
            durationSeries.Name = "Duration";

            // Set the categories (task names) from column A
            chart.NSeries.CategoryData = "A2:A5";

            // Optional visual tweaks
            chart.GapWidth = 50; // 0-500, lower value = thicker bars
            chart.ShowLegend = false;

            // Save the workbook with the new chart
            workbook.Save(outputPath);
            Console.WriteLine($"Gantt chart created and saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
