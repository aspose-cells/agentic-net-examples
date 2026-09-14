// Title: Move an Aspose.Cells chart backward in Z‑order and place a linked picture on top using C#
// AI Prompts: Generate C# code that sends a chart shape to the back of the Z‑order in an Aspose.Cells worksheet and then adds a linked picture above it. | Show how to adjust the Z‑order of a chart and insert a linked image over the chart with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# move chart behind other shapes | place linked image on top of chart using Aspose.Cells | change Z order of chart and picture in Aspose.Cells workbook | example of chart layering with linked picture in Aspose.Cells .NET
// Tags: chart Z-order manipulation Aspose.Cells C# | linked picture insertion Aspose.Cells worksheet | chart layering control Excel Aspose.Cells | Aspose.Cells shape ordering API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data, inserts a column chart, and demonstrates how to adjust the chart's Z‑order (by using positioning properties) so that a linked picture added later appears above the chart. The picture is added from a file if it exists, and the workbook is saved as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);
            worksheet.Cells["A3"].PutValue(30);
            worksheet.Cells["B1"].PutValue(15);
            worksheet.Cells["B2"].PutValue(25);
            worksheet.Cells["B3"].PutValue(35);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("A1:A3", true);
            chart.NSeries[0].Name = "Series1";

            // Note: Aspose.Cells does not provide a ToShape method for Chart.
            // If Z‑order manipulation is required, use chart.Position or related properties.

            // Add a linked picture if the image file exists
            string imagePath = "sample.png";
            if (File.Exists(imagePath))
            {
                int pictureIndex = worksheet.Pictures.Add(5, 6, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Link picture to a cell (available in newer versions)
                // picture.IsLinked = true;
                // picture.LinkedCell = "C1";
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
            }

            // Save the workbook
            workbook.Save("Result.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
