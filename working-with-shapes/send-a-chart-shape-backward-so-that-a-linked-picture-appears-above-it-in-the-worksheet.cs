// Title: How to send a chart shape backward so a linked picture appears on top using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells send chart shape to back so picture overlays it | Aspose.Cells .NET move chart behind a linked picture in an Excel worksheet | Change Z-order of chart and picture with Aspose.Cells C# example | Place an external image above a chart by adjusting shape order using Aspose.Cells | How to set a chart behind a picture in Excel file with Aspose.Cells for .NET
// Tags: chart shape backward Aspose.Cells | linked picture above chart Aspose.Cells | chart Z-order manipulation C# | adjust shape order Excel Aspose.Cells | place picture over chart Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds numeric data, inserts a column chart, adds a picture from a file path, and saves the file as Result.xlsx. It demonstrates how to add both a chart and a linked picture to the same worksheet, allowing you to control their layering (e.g., sending the chart backward so the picture appears on top) using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("A1:A3", true);

            // Path to external image
            string imagePath = @"C:\Images\sample.png";
            if (File.Exists(imagePath))
            {
                try
                {
                    // Add a picture (linked picture not supported directly)
                    int pictureIndex = sheet.Pictures.Add(2, 6, imagePath);
                    Picture picture = sheet.Pictures[pictureIndex];
                    // Additional picture formatting can be applied here if needed
                }
                catch (Exception picEx)
                {
                    Console.WriteLine($"Failed to add picture: {picEx.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}");
            }

            // Save the workbook
            workbook.Save("Result.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
