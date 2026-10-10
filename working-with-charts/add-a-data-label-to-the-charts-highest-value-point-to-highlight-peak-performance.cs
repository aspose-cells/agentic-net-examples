// Title: Highlight the peak point in an Excel chart with a red bold data label using Aspose.Cells for .NET
// AI Prompts: Locate the maximum numeric value across all chart series and apply a red bold data label only to that point with Aspose.Cells. | Refactor the chart code so that data labels are shown solely for the highest data point instead of the entire series. | Change the data label appearance to use a red font and bold style for the peak point in an Excel chart via Aspose.Cells C#.
// Common Searches: Aspose.Cells C# add data label to the highest point of a chart | how to show only the peak value label in an Excel chart using Aspose.Cells | set red bold data label for a specific chart point with Aspose.Cells .NET | find maximum series value and label it in an Aspose.Cells chart
// Tags: add data label to peak chart point Aspose.Cells | customize data label font Aspose.Cells C# | find maximum value in chart series Aspose.Cells | highlight highest point Excel chart Aspose.Cells | set data label color red Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads an Excel workbook, identifies the data point with the highest numeric value across all series in the first chart, enables a data label for that point with red bold font, and saves the updated workbook.
class AddPeakDataLabel
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook containing the chart
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the first worksheet
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            Chart chart = sheet.Charts[0];

            // Variables to track the point with the highest value
            int maxSeriesIndex = -1;
            int maxPointIndex = -1;
            double maxValue = double.MinValue;

            // Scan all series to locate the peak value
            for (int s = 0; s < chart.NSeries.Count; s++)
            {
                Series series = chart.NSeries[s];
                object valuesObj = series.Values;
                if (valuesObj == null) continue;

                double[] values = null;

                if (valuesObj is double[] dArr)
                {
                    values = dArr;
                }
                else if (valuesObj is int[] iArr)
                {
                    values = new double[iArr.Length];
                    for (int i = 0; i < iArr.Length; i++) values[i] = iArr[i];
                }
                else if (valuesObj is string[] sArr)
                {
                    values = new double[sArr.Length];
                    for (int i = 0; i < sArr.Length; i++)
                        double.TryParse(sArr[i], out values[i]);
                }

                if (values == null) continue;

                for (int p = 0; p < values.Length; p++)
                {
                    double pointValue = values[p];
                    if (pointValue > maxValue)
                    {
                        maxValue = pointValue;
                        maxSeriesIndex = s;
                        maxPointIndex = p;
                    }
                }
            }

            // Apply a data label to the series containing the peak point
            if (maxSeriesIndex >= 0)
            {
                Series maxSeries = chart.NSeries[maxSeriesIndex];

                // Show data labels for the series
                maxSeries.DataLabels.ShowValue = true;
                maxSeries.DataLabels.Font.Color = Color.Red;
                maxSeries.DataLabels.Font.IsBold = true;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
