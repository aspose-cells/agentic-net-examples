using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class SetPrimaryYAxisMajorUnit
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                for (int i = 2; i <= 6; i++)
                {
                    worksheet.Cells[$"A{i}"].PutValue("Item " + (i - 1));
                    worksheet.Cells[$"B{i}"].PutValue((i - 1) * 15);
                }

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries.CategoryData = "A2:A6";

                // Disable automatic major unit calculation and set the major unit to 10
                chart.ValueAxis.IsAutomaticMajorUnit = false;
                chart.ValueAxis.MajorUnit = 10;

                // Save the workbook
                string outputPath = "ChartWithMajorUnit.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            SetPrimaryYAxisMajorUnit.Run();
        }
    }
}