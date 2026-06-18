using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class DuplicateChartExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart to duplicate
            Chart originalChart = sheet.Charts[0];

            // Add a new chart with the same type and position as the original chart
            int newChartIndex = sheet.Charts.Add(
                originalChart.Type,
                originalChart.ChartObject.UpperLeftRow,
                originalChart.ChartObject.UpperLeftColumn,
                originalChart.ChartObject.LowerRightRow,
                originalChart.ChartObject.LowerRightColumn);

            Chart duplicatedChart = sheet.Charts[newChartIndex];

            // Copy the data range of the original chart (adjust as needed)
            duplicatedChart.SetChartDataRange("A1:B4", true);

            // Replicate each series from the original chart
            foreach (Series series in originalChart.NSeries)
            {
                // series.Values holds the data range string for the series.
                // The second parameter indicates whether the series is plotted vertically.
                duplicatedChart.NSeries.Add(series.Values, true);
            }

            // Assign a new category axis range to the duplicated chart (adjust as needed)
            duplicatedChart.NSeries.CategoryData = "A2:A5";

            // Optionally, modify other properties (e.g., title) to differentiate the chart.
            duplicatedChart.Title.Text = "Duplicated Chart with New Category Axis";

            // Save the workbook with the duplicated chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}