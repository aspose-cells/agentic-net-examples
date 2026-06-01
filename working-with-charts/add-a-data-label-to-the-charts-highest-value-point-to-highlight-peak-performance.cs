using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class HighlightPeakChartPoint
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");

                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(250);   // Highest value
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(200);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Ensure the chart is calculated so that point coordinates are available
                chart.Calculate();

                // Find the point with the highest Y value in the first series
                Series series = chart.NSeries[0];
                double maxValue = double.MinValue;
                int maxPointIndex = -1;

                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint pt = series.Points[i];
                    // YValue returns object; convert to double safely
                    double y = Convert.ToDouble(pt.YValue);
                    if (y > maxValue)
                    {
                        maxValue = y;
                        maxPointIndex = i;
                    }
                }

                // Highlight the peak point with a data label
                if (maxPointIndex >= 0)
                {
                    ChartPoint peakPoint = series.Points[maxPointIndex];
                    DataLabels lbl = peakPoint.DataLabels;
                    lbl.ShowValue = true;
                    lbl.Text = $"Peak: {maxValue}";
                    lbl.Position = LabelPositionType.Above;
                }

                // Save the workbook
                string outputPath = "ChartPeakHighlight.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}