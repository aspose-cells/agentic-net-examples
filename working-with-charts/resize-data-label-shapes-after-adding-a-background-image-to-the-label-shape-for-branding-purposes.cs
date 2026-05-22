using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

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

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;
            series.DataLabels.Position = LabelPositionType.Center;

            // Path to the branding image that will be used as background for each label
            string brandingImagePath = "branding.png";

            // Verify the branding image exists before attempting to use it
            bool brandingImageExists = File.Exists(brandingImagePath);

            // Iterate through each data point and customize its label
            foreach (ChartPoint point in series.Points)
            {
                // Disable automatic resizing so we can set a fixed size
                point.DataLabels.IsResizeShapeToFitText = false;

                // Set a custom size for the label shape (width and height in pixels)
                point.DataLabels.Width = 80;
                point.DataLabels.Height = 30;

                // If the branding image is available, add it to the worksheet.
                // Aspose.Cells does not provide a direct API to set a picture as the
                // background of a data label, so we only ensure the image is loaded to avoid errors.
                if (brandingImageExists)
                {
                    int picIdx = worksheet.Pictures.Add(0, 0, brandingImagePath);
                    // The picture is added to the worksheet; further customisation can be done here if needed.
                    // For example, you could position the picture manually, but it cannot be assigned to the label.
                    Picture pic = worksheet.Pictures[picIdx];
                    // Optional: hide the picture if it should not be visible.
                    pic.IsHidden = true;
                }

                // Center the text inside the label shape
                point.DataLabels.TextHorizontalAlignment = TextAlignmentType.Center;
                point.DataLabels.TextVerticalAlignment = TextAlignmentType.Center;
            }

            // Save the workbook with the customized chart
            string outputPath = "ChartDataLabelsWithBranding.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Run error: {ex.Message}");
        }
    }
}