// Title: Export a Pie Chart to JPEG with Aspose.Cells C# Chart.ToImage
// Description: Creates a workbook, fills A1:B4 with categories and values, adds a pie chart, and uses Chart.ToImage with ImageType.Jpeg to save the chart as a JPEG file while optionally saving the workbook.
// Keywords: Aspose.Cells | Chart.ToImage | ImageType.Jpeg | C# | .NET | pie chart export | save chart as jpg | Excel chart image | server‑side chart generation | Aspose.Cells example
// Common Searches: Aspose.Cells export pie chart to JPEG C# | Chart.ToImage JPEG example Aspose | How to save an Aspose.Cells chart as JPG | Convert Excel pie chart to JPEG using Aspose | C# generate chart image from workbook
// Developer Intent: Create a JPEG image of a pie chart generated with Aspose.Cells.
// Use Cases: Embed chart JPEGs in PDF reports or PowerPoint slides. | Provide thumbnail previews of Excel charts on a web dashboard. | Automate email notifications that attach chart images. | Store chart snapshots in a content‑management system for archival.
// AI Prompts: Show how to export the same chart as PNG instead of JPEG. | Add a title, legend, and data labels to the pie chart before saving it as JPEG. | Write code to loop through all charts in a workbook and save each as a separate JPEG file. | Explain how to adjust JPEG quality or resolution when using Chart.ToImage.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsPieChartToJpeg
{
    // Creates a workbook, fills A1:B4 with categories and values, adds a pie chart, and uses Chart.ToImage with ImageType.Jpeg to save the chart as a JPEG file while optionally saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 15);
            Chart pieChart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            pieChart.NSeries.Add("B2:B4", true);          // Values
            pieChart.NSeries.CategoryData = "A2:A4";      // Categories

            // Export the chart to a JPEG image using the ImageType enum
            string jpegPath = "PieChartOutput.jpg";
            pieChart.ToImage(jpegPath, ImageType.Jpeg);

            // Optionally save the workbook containing the chart
            workbook.Save("PieChartWorkbook.xlsx");

            Console.WriteLine($"Pie chart exported to JPEG at: {jpegPath}");
        }
    }
}
