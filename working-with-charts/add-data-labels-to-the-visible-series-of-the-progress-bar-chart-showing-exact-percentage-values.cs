// Title: Add percentage data labels to each series of a Progress Bar chart in an existing Excel file using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a .xlsx workbook, finds the first chart, and enables data labels to show both value and percentage for every series with a 0% number format using Aspose.Cells. | Show how to loop through a chart's NSeries collection and set Series.DataLabels.ShowValue, Series.DataLabels.ShowPercentage, and Series.DataLabels.NumberFormat to "0%" in Aspose.Cells. | Provide a complete example that saves the modified workbook to a new file after applying percentage data labels to a progress bar chart.
// Common Searches: aspocells c# add data labels showing percentage to progress bar chart | how to display whole percent values on Excel chart series using Aspose.Cells .NET | enable data labels for each series in an existing chart with Aspose.Cells | format chart data label as 0% without decimal places in Aspose.Cells C#
// Tags: add percentage data labels to chart series Aspose.Cells | configure chart series data labels Aspose.Cells .NET | progress bar chart data label formatting Aspose.Cells | set series data label number format 0% Aspose.Cells | modify existing Excel chart with Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Loads ProgressBarChart.xlsx, accesses the first chart, enables data labels for each series to show values and percentages formatted as whole percentages (0%), and saves the workbook as ProgressBarChart_WithDataLabels.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "ProgressBarChart.xlsx";
            const string outputPath = "ProgressBarChart_WithDataLabels.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook that contains the Progress Bar chart
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index/name as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found on the first worksheet.");
                    return;
                }

                // Get the chart (assumed to be the first chart on the sheet)
                Chart progressBarChart = worksheet.Charts[0];

                // Iterate through all series in the chart
                foreach (Series series in progressBarChart.NSeries)
                {
                    // Enable data labels for the series (HasDataLabel property is not required in newer API versions)
                    // Configure the data label to show value and percentage
                    series.DataLabels.ShowValue = true;
                    series.DataLabels.ShowPercentage = true;

                    // Format the label to display as a percentage with no decimal places
                    series.DataLabels.NumberFormat = "0%";
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors during processing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
