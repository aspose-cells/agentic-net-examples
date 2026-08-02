// Title: Export an Aspose.Cells Chart to JPEG with C# Using the ToImage Method
// Description: The example creates a workbook, fills cells A1:B4 with category and sales data, adds a column chart, and calls chart.ToImage(path, ImageType.Jpeg) to produce a JPEG file of the chart. The workbook is then saved for reference.
// Keywords: Aspose.Cells chart JPEG export | C# ToImage chart image | save Excel chart as image | ImageType.Jpeg Aspose.Cells | export column chart C# | chart snapshot generation | Excel chart to JPEG
// Common Searches: Aspose.Cells export chart as JPEG C# | How to save an Excel chart to image using Aspose.Cells | C# ToImage method example for chart | Convert Excel chart to JPEG with Aspose.Cells | Generate chart image from workbook C#
// Developer Intent: Create a JPEG image of a worksheet chart by invoking Aspose.Cells' ToImage method in C#.
// Use Cases: Embed chart images in web dashboards or reporting portals. | Attach chart snapshots to email newsletters or automated alerts. | Batch‑export localized charts from multi‑language workbooks. | Produce static visuals for documentation or slide decks.
// AI Prompts: Write C# code that builds a pie chart from worksheet data and saves it as a PNG using Aspose.Cells ToImage. | Show how to loop through all charts in a workbook and export each to a separate JPEG file with custom filenames. | Explain how to adjust resolution and compression quality when exporting a chart to JPEG with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartToJpeg
{
    // The example creates a workbook, fills cells A1:B4 with category and sales data, adds a column chart, and calls chart.ToImage(path, ImageType.Jpeg) to produce a JPEG file of the chart. The workbook is then saved for reference.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["B4"].PutValue(1500);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Export the chart to a JPEG image file using ToImage(string, ImageType)
            string imagePath = "ChartOutput.jpeg";
            chart.ToImage(imagePath, ImageType.Jpeg);

            Console.WriteLine($"Chart has been exported to JPEG image at: {imagePath}");

            // Optionally save the workbook for reference
            workbook.Save("ChartWorkbook.xlsx");
        }
    }
}
