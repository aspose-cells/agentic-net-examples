using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    class ResizeDataLabelShapes
    {
        static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook(FileFormatType.Xlsx);
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category 1");
            sheet.Cells["A2"].PutValue("Category 2");
            sheet.Cells["A3"].PutValue("Category 3");
            sheet.Cells["B1"].PutValue(10);
            sheet.Cells["B2"].PutValue(20);
            sheet.Cells["B3"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B1:B3", true);
            chart.NSeries.CategoryData = "A1:A3";

            // Enable data labels
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowValue = true;
            dataLabels.Position = LabelPositionType.Center;

            // Iterate through each point, resize the label shape
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                ChartPoint point = chart.NSeries[0].Points[i];

                // Aspose.Cells does not expose a direct Hyperlink property on ChartPoint.
                // If needed, a hyperlink can be added to the source cell instead.

                // Disable auto‑fit so custom dimensions are applied
                point.DataLabels.IsResizeShapeToFitText = false;

                // Set custom width and height (pixels)
                point.DataLabels.Width = 80;
                point.DataLabels.Height = 30;
            }

            // Save the workbook
            workbook.Save("ResizedDataLabelsWithHyperlink.xlsx");
        }
    }
}