using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class SendChartBackward
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a column chart to the worksheet (correct overload order)
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 2, 2, 12, 12);
                Chart chart = worksheet.Charts[chartIndex];

                // Get the shape of the chart (ChartShape)
                ChartShape chartShape = chart.ChartObject;

                // Path to the picture file
                string picturePath = "linkedImage.png";

                // Ensure the picture file exists before adding it
                if (!File.Exists(picturePath))
                {
                    throw new FileNotFoundException($"Picture file not found: {picturePath}");
                }

                // Add the picture to the worksheet
                int pictureIndex = worksheet.Pictures.Add(5, 5, picturePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Send the chart shape backward in the Z-order so the picture appears on top
                chartShape.ToFrontOrBack(-1);

                // Save the workbook
                string outputPath = "ChartBackwardDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            SendChartBackward.Run();
        }
    }
}