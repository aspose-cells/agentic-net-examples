using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

class ResizeDataLabelShapes
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;
            series.DataLabels.Position = LabelPositionType.Center;

            // Initial calculation to obtain shape dimensions
            chart.Calculate();

            // Store the original chart width (pixels)
            int originalWidth = chart.ChartObject.Width;

            // Increase chart width by 50%
            double newWidthDouble = originalWidth * 1.5;
            int newWidth = (int)newWidthDouble;
            chart.ChartObject.Width = newWidth;

            // Determine scaling factor
            double scaleFactor = (double)newWidth / originalWidth;

            // Resize each data label shape proportionally
            foreach (ChartPoint point in series.Points)
            {
                // Disable automatic resizing so manual dimensions are used
                point.DataLabels.IsResizeShapeToFitText = false;

                // Scale width and height (pixel based)
                int newLabelWidth = (int)(point.DataLabels.WidthPixel * scaleFactor);
                int newLabelHeight = (int)(point.DataLabels.HeightPixel * scaleFactor);

                point.DataLabels.WidthPixel = newLabelWidth;
                point.DataLabels.HeightPixel = newLabelHeight;
            }

            // Re‑calculate to apply the new dimensions
            chart.Calculate();

            // Save the workbook (ensure the directory exists)
            string outputPath = "ResizedDataLabels.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}