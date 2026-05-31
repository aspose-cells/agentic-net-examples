using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the original chart
            Workbook workbook = new Workbook(inputPath);

            // Get the worksheet and the first chart on it (the chart to be cloned)
            Worksheet sourceSheet = workbook.Worksheets[0];
            if (sourceSheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the source worksheet.");
                return;
            }
            Chart sourceChart = sourceSheet.Charts[0];

            // Add a new worksheet that will host the cloned chart
            Worksheet destinationSheet = workbook.Worksheets.Add("ClonedChartSheet");

            // Create a new chart on the destination sheet with the same type and position as the source chart
            int clonedChartIndex = destinationSheet.Charts.Add(
                sourceChart.Type,
                sourceChart.ChartObject.UpperLeftRow,
                sourceChart.ChartObject.UpperLeftColumn,
                sourceChart.ChartObject.LowerRightRow,
                sourceChart.ChartObject.LowerRightColumn);
            Chart clonedChart = destinationSheet.Charts[clonedChartIndex];

            // Retrieve the data range string from the source chart (e.g., "Sheet1!A1:B4")
            string sourceRange = sourceChart.GetChartDataRange();

            // Replace the source sheet name with the destination sheet name in the range string
            string updatedRange = sourceRange.Replace(sourceSheet.Name + "!", destinationSheet.Name + "!");

            // Apply the updated data range to the cloned chart
            clonedChart.SetChartDataRange(updatedRange, true);

            // Save the workbook with the cloned chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}