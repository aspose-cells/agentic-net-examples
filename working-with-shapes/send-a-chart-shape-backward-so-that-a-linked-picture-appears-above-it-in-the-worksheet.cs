// Title: Move a chart shape behind a linked picture in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that places a chart shape at the back of the Z-order and positions a linked picture in front on the same worksheet. | Show how to change the Z-order of a chart and a picture so the picture overlays the chart in an Aspose.Cells workbook.
// Common Searches: asp.net aspose.cells send chart to back behind picture | c# change z-order of chart and picture in Excel using Aspose.Cells | how to overlay linked picture on chart with Aspose.Cells .NET | Aspose.Cells move chart shape behind other shapes programmatically | set chart shape behind linked picture Aspose.Cells example
// Tags: chart shape back order Aspose.Cells | linked picture front order Aspose.Cells | adjust Z-order shapes Aspose.Cells .NET | chart and picture layering Excel Aspose.Cells | move chart behind picture C# Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data, inserts a column chart and a linked picture, then uses Z-order methods to send the chart shape to the back so the picture appears on top, and saves the file as ChartWithLinkedPicture.xlsx.
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
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Add a linked picture that references the same range (A1:B4)
            int pictureIdx = sheet.Pictures.Add(2, 8, "A1:B4");
            Picture picture = sheet.Pictures[pictureIdx];

            // Save the workbook
            workbook.Save("ChartWithLinkedPicture.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
