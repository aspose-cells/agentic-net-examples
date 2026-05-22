using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class RemoveChartFromChartSheet
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a chart sheet (use SheetType.Chart) and give it a name
            int chartSheetIdx = workbook.Worksheets.Add(SheetType.Chart);
            Worksheet chartSheet = workbook.Worksheets[chartSheetIdx];
            chartSheet.Name = "MyChartSheet";

            // Add a sample column chart to the chart sheet
            int chartIdx = chartSheet.Charts.Add(ChartType.Column, 0, 0, 20, 10);
            Chart chart = chartSheet.Charts[chartIdx];

            // Populate data for the chart
            chartSheet.Cells["A1"].PutValue("Category");
            chartSheet.Cells["A2"].PutValue("A");
            chartSheet.Cells["A3"].PutValue("B");
            chartSheet.Cells["B1"].PutValue("Value");
            chartSheet.Cells["B2"].PutValue(10);
            chartSheet.Cells["B3"].PutValue(20);
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // The chart on a chart sheet is stored as a shape.
            // Remove the chart shape (first shape in the collection).
            if (chartSheet.Shapes.Count > 0)
            {
                chartSheet.Shapes.RemoveAt(0);
            }

            // Define output file path
            string outputPath = "ChartSheetWithoutChart.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}