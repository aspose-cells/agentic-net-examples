// Title: Export a Pie Chart to JPEG with Aspose.Cells Chart.ToImage (C#)
// Description: Creates a workbook, adds sample data, builds a pie chart, and saves the chart as a JPEG file using Chart.ToImage with ImageType.Jpeg. The workbook can also be saved for reference.
// Keywords: Aspose.Cells | Chart.ToImage | ImageType.Jpeg | pie chart export | C# | .NET | export Excel chart as JPEG | chart to image example | Aspose.Cells JPEG output | Excel chart image conversion
// Common Searches: Aspose.Cells export pie chart to JPEG C# | Chart.ToImage JPEG example Aspose.Cells | Save Excel chart as JPEG using .NET | How to convert Aspose.Cells chart to image | C# code for exporting chart to JPEG with Aspose
// Developer Intent: Generate a JPEG image file from a pie chart created in an Aspose.Cells workbook.
// Use Cases: Include high‑quality chart images in PDF or web reports without sharing the workbook. | Create thumbnail previews of Excel charts for a dashboard or portal. | Automate batch export of multiple workbook charts to JPEG files for documentation.
// AI Prompts: Write C# code that creates a line chart and saves it as a PNG using Aspose.Cells Chart.ToImage. | Explain how to control JPEG quality and resolution when exporting a chart with Aspose.Cells. | Provide a C# loop that iterates through all charts in a workbook and saves each as a separate JPEG file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartExport
{
    // Creates a workbook, adds sample data, builds a pie chart, and saves the chart as a JPEG file using Chart.ToImage with ImageType.Jpeg. The workbook can also be saved for reference.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
            Chart pieChart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            pieChart.NSeries.Add("B2:B4", true);
            pieChart.NSeries.CategoryData = "A2:A4";

            // Export the chart to a JPEG image using the ImageType enum
            string jpegPath = "PieChart.jpg";
            pieChart.ToImage(jpegPath, ImageType.Jpeg);

            // Optionally save the workbook for reference
            workbook.Save("PieChartWorkbook.xlsx");

            Console.WriteLine($"Pie chart exported to JPEG file: {jpegPath}");
        }
    }
}
