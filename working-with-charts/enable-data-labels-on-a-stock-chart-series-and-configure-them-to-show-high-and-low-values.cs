using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace StockChartDemo
{
    class StockChartDataLabelsDemo
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate header row
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Open");
                sheet.Cells["C1"].PutValue("High");
                sheet.Cells["D1"].PutValue("Low");
                sheet.Cells["E1"].PutValue("Close");

                // Sample data rows
                for (int i = 2; i <= 6; i++)
                {
                    sheet.Cells[i - 1, 0].PutValue($"Day {i - 1}");
                    sheet.Cells[i - 1, 1].PutValue(10 + i); // Open
                    sheet.Cells[i - 1, 2].PutValue(15 + i); // High
                    sheet.Cells[i - 1, 3].PutValue(5 + i);  // Low
                    sheet.Cells[i - 1, 4].PutValue(12 + i); // Close
                }

                // Add a chart sheet
                Worksheet chartSheet = workbook.Worksheets[workbook.Worksheets.Add(SheetType.Chart)];

                // Add a Stock OHLC chart (use StockHighLowClose)
                Chart chart = chartSheet.Charts[chartSheet.Charts.Add(ChartType.StockHighLowClose, 5, 0, 20, 15)];

                // Add series for High values
                Series highSeries = chart.NSeries[chart.NSeries.Add("C2:C6", true)];
                highSeries.Name = "High";

                // Add series for Low values
                Series lowSeries = chart.NSeries[chart.NSeries.Add("D2:D6", true)];
                lowSeries.Name = "Low";

                // Enable data labels and show the values for each series
                highSeries.DataLabels.ShowValue = true;
                lowSeries.DataLabels.ShowValue = true;

                // Optional: set label positions for better readability
                highSeries.DataLabels.Position = LabelPositionType.Above;
                lowSeries.DataLabels.Position = LabelPositionType.Below;

                // Save the workbook
                string outputPath = "StockChartDataLabels.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}