using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class HideSecondSeriesInLineChart
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for two series
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Jan");
                worksheet.Cells["A3"].PutValue("Feb");
                worksheet.Cells["A4"].PutValue("Mar");

                worksheet.Cells["B1"].PutValue("Series1");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                worksheet.Cells["C1"].PutValue("Series2");
                worksheet.Cells["C2"].PutValue(15);
                worksheet.Cells["C3"].PutValue(25);
                worksheet.Cells["C4"].PutValue(35);

                // Add a line chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Add the two series to the chart
                chart.NSeries.Add("B2:B4", true); // First series
                chart.NSeries.Add("C2:C4", true); // Second series
                chart.NSeries.CategoryData = "A2:A4";

                // Hide the second series
                chart.NSeries[1].IsFiltered = true;

                // Save the workbook
                string outputPath = "HideSecondSeriesLineChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            HideSecondSeriesInLineChart.Run();
        }
    }
}