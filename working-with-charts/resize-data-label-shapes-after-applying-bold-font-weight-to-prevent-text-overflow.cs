using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class ResizeDataLabelShapesAfterBoldDemo
    {
        public static void Main()
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            DataLabels labels = chart.NSeries[0].DataLabels;
            labels.ShowValue = true;

            // Apply bold font to the data labels
            labels.Font.IsBold = true;
            labels.Font.Color = Color.Blue;
            labels.Font.Size = 12;

            // Prevent automatic resizing (which would shrink the text)
            labels.IsResizeShapeToFitText = false;

            // Manually increase the label shape size to accommodate the bold text (pixels)
            labels.WidthPixel = 80;
            labels.HeightPixel = 30;

            // Keep the font size constant
            labels.AutoScaleFont = false;

            // Save the workbook
            string outputPath = "ResizeDataLabelShapesAfterBoldDemo.xlsx";
            workbook.Save(outputPath);
        }
    }
}