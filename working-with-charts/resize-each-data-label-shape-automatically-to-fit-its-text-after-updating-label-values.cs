using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class ResizeDataLabelShapesDemo
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Alpha");
            sheet.Cells["A3"].PutValue("Beta");
            sheet.Cells["A4"].PutValue("Gamma");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["B3"].PutValue(4567);
            sheet.Cells["B4"].PutValue(89);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;
            series.DataLabels.Position = LabelPositionType.Center;

            // Update each data label's text and enable auto‑resize to fit the text
            foreach (ChartPoint point in series.Points)
            {
                // Prepend a custom prefix to the label text
                point.DataLabels.Text = $"Val: {point.YValue}";

                // Enable automatic resizing of the label shape to fit its text
                point.DataLabels.IsResizeShapeToFitText = true;
            }

            // Define output file path
            string outputPath = "ResizeDataLabelShapesDemo.xlsx";

            // Save the workbook (overwrite if exists)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
        }
    }
}