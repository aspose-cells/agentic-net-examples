using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class BarChartWithSecondaryAxis
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                worksheet.Cells["A1"].PutValue("Month");
                worksheet.Cells["A2"].PutValue("Jan");
                worksheet.Cells["A3"].PutValue("Feb");
                worksheet.Cells["A4"].PutValue("Mar");

                worksheet.Cells["B1"].PutValue("Revenue");
                worksheet.Cells["B2"].PutValue(120000);
                worksheet.Cells["B3"].PutValue(150000);
                worksheet.Cells["B4"].PutValue(130000);

                worksheet.Cells["C1"].PutValue("Units Sold");
                worksheet.Cells["C2"].PutValue(3000);
                worksheet.Cells["C3"].PutValue(3500);
                worksheet.Cells["C4"].PutValue(3200);

                // Add a 2‑D clustered bar chart
                int chartIndex = worksheet.Charts.Add(ChartType.Bar, 6, 0, 20, 12);
                Chart chart = worksheet.Charts[chartIndex];

                // Add data series
                chart.NSeries.Add("C2:C4", true); // Units Sold (primary axis)
                chart.NSeries.Add("B2:B4", true); // Revenue (secondary axis)

                // Set category (X) axis data
                chart.NSeries.CategoryData = "A2:A4";

                // Plot Revenue series on secondary axis (second series, index 1)
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Customize secondary axis
                Axis secondaryAxis = chart.SecondValueAxis;
                secondaryAxis.Title.Text = "Revenue (USD)";
                secondaryAxis.MinValue = 0;
                secondaryAxis.MaxValue = 200000;
                secondaryAxis.MajorUnit = 50000;

                // Customize primary axis title
                chart.ValueAxis.Title.Text = "Units Sold";

                // Save the workbook
                string outputPath = "BarChartWithSecondaryAxis.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            BarChartWithSecondaryAxis.Run();
        }
    }
}