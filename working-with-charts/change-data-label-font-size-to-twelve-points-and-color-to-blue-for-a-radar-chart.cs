using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class RadarChartDataLabelFormatting
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the radar chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Cat1");
                worksheet.Cells["A3"].PutValue("Cat2");
                worksheet.Cells["A4"].PutValue("Cat3");
                worksheet.Cells["A5"].PutValue("Cat4");

                worksheet.Cells["B1"].PutValue("Series1");
                worksheet.Cells["B2"].PutValue(4);
                worksheet.Cells["B3"].PutValue(2);
                worksheet.Cells["B4"].PutValue(5);
                worksheet.Cells["B5"].PutValue(3);

                // Add a radar chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Radar, 5, 0, 20, 12);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Access the first series
                Series series = chart.NSeries[0];

                // Enable data labels and format them
                series.DataLabels.ShowValue = true;
                series.DataLabels.Font.Size = 12;
                series.DataLabels.Font.Color = Color.Blue;
                series.DataLabels.ApplyFont();

                // Define output file path
                string outputPath = "RadarChartDataLabelsFormatted.xlsx";

                // Save the workbook (overwrite if it already exists)
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any runtime errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            RadarChartDataLabelFormatting.Run();
        }
    }
}