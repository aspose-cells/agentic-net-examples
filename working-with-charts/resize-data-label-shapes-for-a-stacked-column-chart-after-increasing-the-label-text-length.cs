using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelResizeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a stacked column chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");

                // Series 1
                sheet.Cells["B1"].PutValue("Product A");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(40);
                sheet.Cells["B4"].PutValue(20);

                // Series 2
                sheet.Cells["C1"].PutValue("Product B");
                sheet.Cells["C2"].PutValue(20);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a stacked column chart (use ColumnStacked enum)
                int chartIdx = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Set data range for both series
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries[0].Name = "Product A";
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries[1].Name = "Product B";

                // Set category (X) axis data
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for both series
                foreach (Series ser in chart.NSeries)
                {
                    ser.DataLabels.ShowValue = true;
                    ser.DataLabels.Position = LabelPositionType.Center;
                }

                // Increase the label text length for each point (e.g., prepend a long description)
                foreach (ChartPoint point in chart.NSeries[0].Points)
                {
                    point.DataLabels.Text = "VeryLongLabel_" + point.YValue;
                    // Disable auto‑resize so we can set a custom size
                    point.DataLabels.IsResizeShapeToFitText = false;
                    // Set a width that is smaller than the text would normally need
                    point.DataLabels.Width = 80;   // pixels
                    point.DataLabels.Height = 30;  // pixels
                }

                // For the second series, let the shape auto‑fit to the longer text
                foreach (ChartPoint point in chart.NSeries[1].Points)
                {
                    point.DataLabels.Text = "AnotherVeryLongLabel_" + point.YValue;
                    // Enable auto‑fit so the shape expands to contain the text
                    point.DataLabels.IsResizeShapeToFitText = true;
                }

                // Recalculate the chart to apply changes
                chart.Calculate();

                // Save the workbook
                string outputPath = "StackedColumnDataLabelResize.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}