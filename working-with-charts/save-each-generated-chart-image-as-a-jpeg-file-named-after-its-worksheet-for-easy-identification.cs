// Title: Save worksheet charts as JPEG files named after their sheets using Aspose.Cells for .NET
// Description: A C# sample that builds a workbook, inserts sample data and a column chart on each sheet, then loops through the sheets to export every chart as a JPEG image whose filename mirrors the sheet name. The workbook can be saved afterwards.
// Keywords: Aspose.Cells C# chart export | Chart.ToImage JPEG | save chart per worksheet | export workbook charts .NET | image naming by sheet | column chart image Aspose | automate chart extraction | Aspose.Cells image generation
// Common Searches: Aspose.Cells export chart to JPEG per sheet | C# save each worksheet chart as image | Chart.ToImage example Aspose.Cells | how to name chart images after worksheet | batch export charts from Excel using Aspose
// Developer Intent: Create JPEG images for all charts in a workbook, using each worksheet's name as the image file name.
// Use Cases: Produce individual chart pictures for embedding in web reports or email summaries. | Generate thumbnail previews of sheet charts for a dashboard overview. | Automate extraction of chart graphics for documentation, presentations, or archival purposes.
// AI Prompts: Write C# code that iterates through every worksheet in an Aspose.Cells workbook and saves each chart as a PNG file named after the worksheet, placing the images in a designated folder. | Provide an example that exports charts as high‑resolution JPEGs, includes error handling for sheets without charts, and returns a list of the generated file paths. | Show how to modify the sample to add the worksheet index to the JPEG filename and set a custom image quality parameter.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// A C# sample that builds a workbook, inserts sample data and a column chart on each sheet, then loops through the sheets to export every chart as a JPEG image whose filename mirrors the sheet name. The workbook can be saved afterwards.
class SaveChartImages
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data and a chart to each worksheet
        for (int w = 0; w < workbook.Worksheets.Count; w++)
        {
            Worksheet ws = workbook.Worksheets[w];
            ws.Name = $"Sheet{w + 1}";

            ws.Cells["A1"].PutValue("Category");
            ws.Cells["A2"].PutValue("Apple");
            ws.Cells["A3"].PutValue("Orange");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["B2"].PutValue(10 + w * 5);
            ws.Cells["B3"].PutValue(20 + w * 5);

            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 8);
            Chart chart = ws.Charts[chartIdx];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";
        }

        // Save each chart as a JPEG file named after its worksheet
        foreach (Worksheet ws in workbook.Worksheets)
        {
            for (int i = 0; i < ws.Charts.Count; i++)
            {
                Chart chart = ws.Charts[i];
                string imageFile = $"{ws.Name}.jpg";
                chart.ToImage(imageFile, ImageType.Jpeg);
            }
        }

        // Save the workbook (optional)
        workbook.Save("WorkbookWithCharts.xlsx");
    }
}
