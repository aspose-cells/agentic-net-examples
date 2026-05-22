using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class VerifyChartParentWorksheet
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (creation rule)
                Workbook workbook = new Workbook();

                // Access the first worksheet and set a known name
                Worksheet worksheet = workbook.Worksheets[0];
                string expectedWorksheetName = "DataSheet";
                worksheet.Name = expectedWorksheetName;

                // Add sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Retrieve the worksheet that contains the chart
                Worksheet chartParentWorksheet = chart.Worksheet;

                // Verify that the retrieved worksheet name matches the expected name
                bool namesMatch = string.Equals(chartParentWorksheet.Name, expectedWorksheetName, StringComparison.Ordinal);
                Console.WriteLine($"Expected worksheet name: {expectedWorksheetName}");
                Console.WriteLine($"Chart's parent worksheet name: {chartParentWorksheet.Name}");
                Console.WriteLine($"Names match: {namesMatch}");

                // Save the workbook (save rule)
                string outputPath = "VerifyChartParentWorksheet_out.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            VerifyChartParentWorksheet.Run();
        }
    }
}