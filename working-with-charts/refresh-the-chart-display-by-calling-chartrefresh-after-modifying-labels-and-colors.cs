using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace RefreshChartDemoApp
{
    public class RefreshChartDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
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
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series and change their font color
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;
                series.DataLabels.Font.Color = Color.Green;
                series.DataLabels.ApplyFont();

                // Optional: change series colors using a valid palette (commented out because the enum value may differ across versions)
                // chart.NSeries.ChangeColors(ChartColorPaletteType.Default);

                // Save the workbook; the chart is automatically updated
                workbook.Save("RefreshChartDemo.xlsx");
            }
            catch (Exception ex)
            {
                // Log any runtime errors that occur during chart creation or saving
                Console.WriteLine($"Run error: {ex.Message}");
                throw; // Re‑throw to allow outer handler to capture if needed
            }
        }
    }
}