// Title: Add an external PNG picture as a shape positioned over a chart's plot area and save the workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# place picture shape at chart plot area coordinates | retrieve plot area X Y width height from a chart with Aspose.Cells .NET | insert external image behind a column chart and align to plot area using Aspose.Cells | C# Aspose.Cells add picture to worksheet and set ZOrder behind chart | save Excel workbook with chart and overlaid picture shape using Aspose.Cells
// Tags: add picture shape to worksheet using Aspose.Cells | chart plot area dimensions Aspose.Cells | set picture shape bounds from chart coordinates | place external image behind Excel chart Aspose.Cells | save workbook with embedded picture Aspose.Cells Xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExternalShapeDemo
{
    // The example creates a workbook, fills it with sample data, adds a column chart, reads the chart's PlotArea X, Y, Width, and Height, generates a temporary PNG image, inserts the image as a picture shape, aligns the shape's position and size to match the PlotArea, sends the picture behind the chart, and saves the file as ChartWithExternalShape.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Prepare a simple workbook with sample data
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample data for the chart
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("A");
                cells["A3"].PutValue("B");
                cells["A4"].PutValue("C");
                cells["B2"].PutValue(10);
                cells["B3"].PutValue(20);
                cells["B4"].PutValue(30);

                // -----------------------------------------------------------------
                // 2. Add a column chart to the worksheet
                // -----------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // -----------------------------------------------------------------
                // 3. Retrieve the PlotArea position (X, Y, Width, Height)
                // -----------------------------------------------------------------
                PlotArea plotArea = chart.PlotArea;
                double plotX = plotArea.X;
                double plotY = plotArea.Y;
                double plotWidth = plotArea.Width;
                double plotHeight = plotArea.Height;

                Console.WriteLine($"PlotArea Position - X:{plotX}, Y:{plotY}, Width:{plotWidth}, Height:{plotHeight}");

                // -----------------------------------------------------------------
                // 4. Create a temporary image file to use as the external shape
                // -----------------------------------------------------------------
                string imagePath = Path.Combine(Environment.CurrentDirectory, "sample.png");
                CreateSampleImage(imagePath);

                // -----------------------------------------------------------------
                // 5. Add the picture (external shape) to the worksheet
                // -----------------------------------------------------------------
                if (!File.Exists(imagePath))
                    throw new FileNotFoundException("Image file not found.", imagePath);

                // AddPicture now returns a Picture object (Aspose.Cells 23+)
                Picture pictureShape;
                using (FileStream imgStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                {
                    pictureShape = sheet.Shapes.AddPicture(0, 0, 0, 0, imgStream);
                }

                // Set the picture's position and size to match the PlotArea
                pictureShape.X = (int)plotX;
                pictureShape.Y = (int)plotY;
                pictureShape.Width = (int)plotWidth;
                pictureShape.Height = (int)plotHeight;

                // Optional: send the picture behind the chart.
                pictureShape.ZOrderPosition = 0; // backmost

                // -----------------------------------------------------------------
                // 6. Save the workbook
                // -----------------------------------------------------------------
                string outputPath = Path.Combine(Environment.CurrentDirectory, "ChartWithExternalShape.xlsx");
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        /// <param name="filePath">Full path where the image will be saved.</param>
        private static void CreateSampleImage(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                    return;

                // Base64-encoded 1x1 white PNG.
                const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/6V4AAAAASUVORK5CYII=";
                byte[] pngBytes = Convert.FromBase64String(base64Png);
                File.WriteAllBytes(filePath, pngBytes);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create sample image: {ex.Message}");
                throw;
            }
        }
    }
}
