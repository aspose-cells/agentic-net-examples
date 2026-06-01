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

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the first chart that is linked to a PivotTable
            Chart pivotChart = null;
            foreach (Chart chart in worksheet.Charts)
            {
                if (!string.IsNullOrEmpty(chart.PivotSource))
                {
                    pivotChart = chart;
                    break;
                }
            }

            if (pivotChart != null)
            {
                // Set the legend position to the bottom of the chart
                pivotChart.Legend.Position = LegendPositionType.Bottom;

                // Ensure drop zones are visible (optional)
                PivotOptions pivotOptions = pivotChart.PivotOptions;
                if (pivotOptions != null)
                {
                    pivotOptions.DropZonesVisible = true;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}