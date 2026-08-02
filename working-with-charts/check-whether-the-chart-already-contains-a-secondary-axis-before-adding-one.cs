using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class CheckAndAddSecondaryAxis
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

                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["B4"].PutValue(300);

                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["C2"].PutValue(5000);
                sheet.Cells["C3"].PutValue(3000);
                sheet.Cells["C4"].PutValue(1000);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Add two series to the chart
                chart.NSeries.Add("B2:B4", true); // Series 1
                chart.NSeries.Add("C2:C4", true); // Series 2
                chart.NSeries.CategoryData = "A2:A4";

                // Check if a secondary value axis already exists
                bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false);

                if (!hasSecondaryValueAxis)
                {
                    // No secondary axis – enable it by plotting the second series on it
                    chart.NSeries[1].PlotOnSecondAxis = true;

                    // Customize the newly created secondary axis
                    Axis secondValueAxis = chart.SecondValueAxis;
                    secondValueAxis.Title.Text = "Secondary Axis";
                    secondValueAxis.IsVisible = true;
                }

                // Define output file path
                string outputPath = "CheckAndAddSecondaryAxis.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            CheckAndAddSecondaryAxis.Run();
        }
    }
}