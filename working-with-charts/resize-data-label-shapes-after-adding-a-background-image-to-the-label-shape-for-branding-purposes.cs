// Title: Resize Excel chart data label shapes to fixed size after adding a branding image with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to create a column chart, enable data labels, disable automatic shape resizing, and set each label's width to 120 px and height to 50 px. | Show how to insert a PNG branding image into a worksheet and reference it while customizing chart data labels in Aspose.Cells, keeping label dimensions unchanged. | Adapt the example to apply a picture fill to data labels (if supported) while preserving the custom label size settings in Aspose.Cells for .NET.
// Common Searches: how to set fixed width and height for chart data labels in Aspose.Cells C# | add background picture to Excel chart data labels using Aspose.Cells .NET | disable data label auto resize Aspose.Cells chart series | customize size of data label shapes programmatically with Aspose.Cells | Aspose.Cells chart label branding image not resizing
// Tags: Aspose.Cells chart data label size | fixed dimensions for chart data labels .NET | branding image on chart data labels Aspose.Cells | disable automatic data label resize Aspose.Cells | C# set chart label width and height

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsDataLabelResizeDemo
{
    // Demonstrates creating a workbook with a column chart, enabling centered data labels, inserting a branding PNG into the worksheet, disabling automatic label resizing, and setting each data label shape to a fixed width of 120 px and height of 50 px before saving the file as DataLabelResizeWithBranding.xlsx.
    class Program
    {
        static void Main()
        {
            try
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
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.Center;

                // Path to the branding image (ensure the file exists at this location)
                string imagePath = "branding.png";

                // Verify that the image file exists before using it
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}. Skipping picture fill.");
                }

                // Iterate through each point and customize its data label
                foreach (ChartPoint point in series.Points)
                {
                    // If the image exists, add it to the worksheet (the picture itself is not applied to the label
                    // because Aspose.Cells does not expose a direct Fill property on DataLabels)
                    if (File.Exists(imagePath))
                    {
                        int picIdx = sheet.Pictures.Add(0, 0, imagePath);
                        // The picture is added to the worksheet; further customisation can be done if needed.
                        // Note: Direct picture fill for data labels is not supported via the current API.
                        Picture pic = sheet.Pictures[picIdx];
                        // Placeholder for any future operations with 'pic'.
                    }

                    // Disable automatic resizing so we can set a fixed size
                    point.DataLabels.IsResizeShapeToFitText = false;

                    // Set custom dimensions for the label shape (in pixels)
                    point.DataLabels.Width = 120;
                    point.DataLabels.Height = 50;
                }

                // Save the workbook
                string outputPath = "DataLabelResizeWithBranding.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
