using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartConversion
{
    class Program
    {
        static void Main()
        {
            const string inputFile = "InputWithPieChart.xlsx";
            const string outputFile = "OutputWithDoughnutChart.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            try
            {
                // Load the workbook containing the pie chart
                Workbook workbook = new Workbook(inputFile);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet has at least one chart
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Get the first chart (assumed to be a pie chart)
                Chart chart = sheet.Charts[0];

                // Change the chart type from Pie to Doughnut
                chart.Type = ChartType.Doughnut;

                // Example: adjust doughnut-specific property (optional)
                // chart.NSeries[0].DoughnutHoleSize = 30; // 30% hole size

                // Save the modified workbook
                workbook.Save(outputFile);

                Console.WriteLine($"Chart type changed to Doughnut and workbook saved as '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}