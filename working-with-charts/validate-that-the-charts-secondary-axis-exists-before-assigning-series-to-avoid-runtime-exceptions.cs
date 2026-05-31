using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryAxisValidation
{
    public class ValidateSecondaryAxis
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Primary Series");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                sheet.Cells["C1"].PutValue("Secondary Series");
                sheet.Cells["C2"].PutValue(500);
                sheet.Cells["C3"].PutValue(300);
                sheet.Cells["C4"].PutValue(100);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Set data for the primary series
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Add a second series (intended for secondary axis)
                chart.NSeries.Add("C2:C4", true);

                // Validate that the chart supports a secondary value axis before assigning
                bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false);
                if (hasSecondaryValueAxis)
                {
                    // Plot the second series on the secondary axis
                    chart.NSeries[1].PlotOnSecondAxis = true;

                    // Optionally customize the secondary axis
                    chart.SecondValueAxis.IsVisible = true;
                    chart.SecondValueAxis.Title.Text = "Secondary Axis";
                }
                else
                {
                    Console.WriteLine("The selected chart type does not support a secondary value axis. " +
                                      "The second series will remain on the primary axis.");
                }

                // Save the workbook
                string outputPath = "ValidatedSecondaryAxisChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateSecondaryAxis.Run();
        }
    }
}