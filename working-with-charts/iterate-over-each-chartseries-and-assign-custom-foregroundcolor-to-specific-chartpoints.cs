using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ChartSeriesPointsForegroundColor
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["A5"].PutValue("D");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);
                sheet.Cells["C5"].PutValue(45);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for both series
                chart.NSeries.Add("B2:C5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Iterate over each series in the chart
                for (int s = 0; s < chart.NSeries.Count; s++)
                {
                    Series series = chart.NSeries[s];
                    ChartPointCollection points = series.Points;

                    // Iterate over each point in the current series
                    for (int p = 0; p < points.Count; p++)
                    {
                        ChartPoint point = points[p];

                        // Assign a different foreground color based on point index
                        if (p % 3 == 0)
                        {
                            point.Area.ForegroundColor = Color.FromArgb(255, 99, 71); // Tomato
                        }
                        else if (p % 3 == 1)
                        {
                            point.Area.ForegroundColor = Color.FromArgb(60, 179, 113); // MediumSeaGreen
                        }
                        else
                        {
                            point.Area.ForegroundColor = Color.FromArgb(30, 144, 255); // DodgerBlue
                        }

                        // Ensure the formatting type is set to Custom so the color is applied
                        point.Area.Formatting = FormattingType.Custom;
                    }
                }

                // Save the workbook with the customized chart
                string outputPath = "ChartSeriesPointsForegroundColor.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ChartSeriesPointsForegroundColor.Run();
        }
    }
}