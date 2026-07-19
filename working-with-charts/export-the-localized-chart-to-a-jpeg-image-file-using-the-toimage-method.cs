// Title: Export an Excel chart to JPEG with Aspose.Cells C# ToImage method
// Description: Shows how to create a workbook, populate it with sample data, add a column chart, and write the chart to a JPEG file using Chart.ToImage (ImageType.Jpeg) in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | Export chart JPEG | Chart.ToImage | ImageType.Jpeg | Excel chart as image | Aspose.Cells .NET example | save chart to file | chart image conversion
// Common Searches: Aspose.Cells export chart to JPEG C# | How to save Excel chart as JPEG using Aspose.Cells | Chart.ToImage method example .NET | Convert Excel chart to image with Aspose.Cells | C# Aspose.Cells chart image export
// Developer Intent: Create a JPEG image file from a chart that was built inside an Aspose.Cells workbook.
// Use Cases: Generate a chart image for inclusion in web reports or email newsletters. | Provide localized chart graphics (e.g., language‑specific titles) as static JPEG assets. | Archive chart visuals while keeping the original workbook for further data analysis.
// AI Prompts: Write C# code that creates a pie chart from worksheet data and saves it as a PNG using Aspose.Cells. | Show how to set a localized title on a chart and export it to a high‑resolution JPEG with Aspose.Cells. | Provide a loop that exports every chart in a workbook to separate JPEG files using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, populate it with sample data, add a column chart, and write the chart to a JPEG file using Chart.ToImage (ImageType.Jpeg) in Aspose.Cells for .NET.
class ExportChartToJpeg
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["A4"].PutValue("Banana");

        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["B2"].PutValue(1200);
        worksheet.Cells["B3"].PutValue(800);
        worksheet.Cells["B4"].PutValue(1500);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Export the chart to a JPEG image file
        chart.ToImage("ChartOutput.jpg", ImageType.Jpeg);

        // Save the workbook (optional, for reference)
        workbook.Save("ChartWorkbook.xlsx");
    }
}
