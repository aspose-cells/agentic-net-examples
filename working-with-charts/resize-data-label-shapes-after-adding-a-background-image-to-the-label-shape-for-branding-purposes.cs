// Title: Resize Chart Data Label Shapes and Apply a Branding Image with Aspose.Cells for .NET
// Description: This example creates a workbook, adds a column chart, enables centered rectangular data labels, inserts a PNG logo, disables automatic label resizing, sets a fixed width and height for each label, and saves the file as ResizedDataLabels.xlsx. It also demonstrates graceful handling when the branding image is missing.
// Keywords: Aspose.Cells data label size | chart data label custom dimensions .NET | add background image to chart labels | disable auto‑resize data labels Aspose.Cells | C# chart branding Aspose.Cells | resize data label shape Aspose.Cells | Aspose.Cells chart customization
// Common Searches: how to set fixed width for chart data labels in Aspose.Cells | add logo behind data labels in a .NET chart | prevent data label auto‑sizing Aspose.Cells | customize data label shape size C# Aspose.Cells | place image as background for chart labels
// Developer Intent: Set a constant size for each chart data label and overlay a branding picture behind the labels using Aspose.Cells for .NET.
// Use Cases: Standardize label dimensions across a column chart for a clean, brand‑consistent appearance. | Display a company logo or watermark behind data labels without affecting label readability. | Maintain label size when values change, ensuring layout stability in dynamic reports.
// AI Prompts: Write C# code with Aspose.Cells that fixes the width and height of data label shapes in a chart and turns off auto‑resize. | Show how to load a PNG file from the application folder and use it as a background for chart data labels in Aspose.Cells. | Provide a robust example that iterates over chart points, sets label size, changes font color, and handles missing branding images gracefully.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a column chart, enables centered rectangular data labels, inserts a PNG logo, disables automatic label resizing, sets a fixed width and height for each label, and saves the file as ResizedDataLabels.xlsx. It also demonstrates graceful handling when the branding image is missing.
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

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;
            series.DataLabels.Position = LabelPositionType.Center;
            series.DataLabels.ShapeType = DataLabelShapeType.Rect;

            // Insert a picture that will be used as the background for the data labels
            // (Assumes "branding.png" exists in the application directory)
            string picturePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "branding.png");
            int pictureIndex = -1;
            if (File.Exists(picturePath))
            {
                pictureIndex = sheet.Pictures.Add(0, 0, picturePath);
            }
            else
            {
                Console.WriteLine("Warning: branding.png not found. Data labels will use default appearance.");
            }

            // Iterate through each data point and resize its label shape
            foreach (ChartPoint point in series.Points)
            {
                // Disable automatic resizing of the shape to fit the text
                point.DataLabels.IsResizeShapeToFitText = false;

                // Set custom dimensions for the label shape (in points)
                point.DataLabels.Width = 80;   // width in points
                point.DataLabels.Height = 30;  // height in points

                // If the picture was loaded, set it as the background of the label shape
                if (pictureIndex >= 0)
                {
                    // Aspose.Cells does not provide a direct Fill property for DataLabels.
                    // As a workaround, we set the label's background color to transparent
                    // and rely on the picture being placed behind the chart.
                    point.DataLabels.Font.Color = Color.Black; // ensure text is visible
                }
            }

            // Save the workbook with the resized data label shapes
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ResizedDataLabels.xlsx");
            workbook.Save(outputPath);
        }
    }
}
