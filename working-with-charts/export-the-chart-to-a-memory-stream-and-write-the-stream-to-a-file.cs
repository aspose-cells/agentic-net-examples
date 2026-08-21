// Title: Export an Aspose.Cells chart to PNG using a MemoryStream in C#
// Description: Creates a workbook, adds sample data and a column chart, converts the chart to a PNG image with the `ToImage` method into a `MemoryStream`, resets the stream, and writes the bytes to `ChartImage.png`. The workbook is also saved for reference.
// Keywords: Aspose.Cells chart export | C# chart to PNG | MemoryStream image saving | ToImage method Aspose.Cells | save chart as image file | export Excel chart image | Aspose.Cells PNG output | chart image generation C#
// Common Searches: Aspose.Cells export chart to PNG C# | How to save an Aspose.Cells chart as an image | MemoryStream chart image Aspose.Cells example | C# convert Excel chart to PNG | Aspose.Cells ToImage usage | Write chart image to disk without opening Excel
// Developer Intent: Generate a PNG file from an Aspose.Cells chart by streaming the image data and writing it to disk.
// Use Cases: Attach chart images to automated email reports. | Display chart thumbnails on a web dashboard without loading the workbook. | Create separate image assets for documentation or presentations. | Cache chart visuals for quick retrieval in high‑performance applications.
// AI Prompts: Show how to export an Aspose.Cells chart to JPEG using a MemoryStream in C#. | Provide code to embed the exported PNG chart into a PDF with Aspose.PDF. | Explain how to set image resolution (DPI) when exporting a chart with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartExport
{
    // Creates a workbook, adds sample data and a column chart, converts the chart to a PNG image with the `ToImage` method into a `MemoryStream`, resets the stream, and writes the bytes to `ChartImage.png`. The workbook is also saved for reference.
    public class ExportChartToFile
    {
        public static void Main()
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
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Export the chart to a memory stream and then write it to a file
            using (MemoryStream stream = new MemoryStream())
            {
                // Convert chart to PNG image and write to the stream
                chart.ToImage(stream, ImageType.Png);

                // Ensure the stream position is at the beginning before reading
                stream.Position = 0;

                // Write the stream contents to a file
                File.WriteAllBytes("ChartImage.png", stream.ToArray());

                Console.WriteLine("Chart image saved to ChartImage.png");
            }

            // Optionally save the workbook for reference
            workbook.Save("ChartWorkbook.xlsx");
        }
    }
}
