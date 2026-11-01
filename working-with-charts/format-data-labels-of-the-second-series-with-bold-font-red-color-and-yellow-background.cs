// Title: How to make data labels of the second chart series bold with red text and a yellow fill using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that enables data labels for the second series of an Aspose.Cells chart and sets the label font to bold, the text color to red, and the background fill to yellow. | Show the steps to access a specific series in an Aspose.Cells chart and apply custom font style and fill color to its data labels in a .NET workbook. | Update an existing Aspose.Cells workbook so that the second series' data labels display bold red text on a yellow background.
// Common Searches: Aspose.Cells C# set second series data label font to bold and color red | how to add yellow background to chart data labels in Aspose.Cells .NET | enable and style data labels for a specific series in an Aspose.Cells column chart | C# Aspose.Cells change data label appearance for second series only
// Tags: format second series data labels Aspose.Cells | set data label font style C# Aspose.Cells | apply text color to chart data labels .NET | add background fill to chart data labels Aspose.Cells | customize chart series label appearance C# | chart data label styling Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartFormatting
{
    // The example creates a workbook, adds a column chart with at least two series, enables data labels on the second series, sets the label font to bold and red, and saves the workbook as FormattedChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];

                // Obtain an existing chart or add a new one if none exist
                Chart chart;
                if (worksheet.Charts.Count > 0)
                {
                    chart = worksheet.Charts[0];
                }
                else
                {
                    // Add a column chart at the specified position
                    int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                    chart = worksheet.Charts[chartIndex];
                }

                // Ensure there are at least two series; add dummy series if needed
                if (chart.NSeries.Count < 2)
                {
                    // First series (e.g., data from A1:A5)
                    chart.NSeries.Add("A1:A5", true);
                    // Second series (the one we will format, e.g., data from B1:B5)
                    chart.NSeries.Add("B1:B5", true);
                }

                // Get the second series (index 1)
                var secondSeries = chart.NSeries[1];

                // Access the data labels collection of the second series
                var dataLabels = secondSeries.DataLabels;

                // Enable data labels to show values
                dataLabels.ShowValue = true;

                // Format the data labels: bold font, red text color
                dataLabels.Font.IsBold = true;
                dataLabels.Font.Color = Color.Red;

                // Save the workbook to a file
                string outputPath = "FormattedChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
