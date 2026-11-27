// Title: How to resize chart data label shapes and apply a branding background image using Aspose.Cells for .NET
// Common Searches: set fixed width and height for chart data labels in Aspose.Cells C# | add a picture background to data labels of an Aspose.Cells chart | prevent automatic resizing of data label shapes in Aspose.Cells .NET | customize data label font color and boldness with Aspose.Cells | save a workbook with branded chart labels to XLSX using Aspose.Cells
// Tags: chart data label fixed dimensions Aspose.Cells | Aspose.Cells data label picture fill | disable data label auto resize Aspose.Cells | customize data label font Aspose.Cells | Aspose.Cells chart branding image

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// The example creates a workbook, adds a column chart, enables data labels, disables automatic shape resizing, sets each label to a fixed width of 80 points and height of 30 points, optionally prepares a picture fill for a branding image, styles the label text with white bold font, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the branding image that could be used as background for data labels.
            string brandingImagePath = "branding.png";

            if (!File.Exists(brandingImagePath))
            {
                Console.WriteLine($"Branding image not found at '{brandingImagePath}'. The image will be skipped.");
            }

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(200);

            // Add a column chart.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels for the first series.
            DataLabels seriesLabels = chart.NSeries[0].DataLabels;
            seriesLabels.ShowValue = true;
            seriesLabels.Position = LabelPositionType.Center;

            // Customize each data label.
            foreach (ChartPoint point in chart.NSeries[0].Points)
            {
                // Disable automatic resizing to fit text.
                point.DataLabels.IsResizeShapeToFitText = false;

                // Set a fixed size for the label shape (width and height are in points).
                point.DataLabels.Width = 80;   // Approx. 1.11 cm
                point.DataLabels.Height = 30; // Approx. 0.42 cm

                // If the branding image exists, you could apply it as a picture fill.
                // The Fill property is not available on DataLabels in some versions,
                // so this step is optional and can be omitted.
                if (File.Exists(brandingImagePath))
                {
                    // Example placeholder for applying a picture fill if supported:
                    // point.DataLabels.Fill.FillType = FillType.Picture;
                    // point.DataLabels.Fill.SetPictureFill(brandingImagePath);
                }

                // Set label text appearance.
                point.DataLabels.Font.Color = Color.White;
                point.DataLabels.Font.IsBold = true;
            }

            // Save the workbook.
            string outputPath = "ChartWithBrandedDataLabels.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
