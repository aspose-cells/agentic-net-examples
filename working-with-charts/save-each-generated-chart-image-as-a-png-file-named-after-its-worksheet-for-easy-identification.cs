// Title: Save Excel worksheet charts as PNG files named after their sheets using Aspose.Cells for .NET
// Description: Demonstrates how to loop through all worksheets in a workbook, create a column chart on each sheet, and export each chart to a PNG image whose filename matches the worksheet name with Aspose.Cells Chart.ToImage. The workbook is then saved as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | export chart as PNG | Chart.ToImage | save chart image | worksheet name file | batch chart export | Excel chart image extraction | automate chart saving
// Common Searches: Aspose.Cells export chart to PNG C# | save Excel chart image with worksheet name | Chart.ToImage example Aspose.Cells | batch export charts from workbook .NET | C# code to save each sheet chart as PNG
// Developer Intent: Automatically generate a PNG image for every chart in a workbook, naming each file after its corresponding worksheet.
// Use Cases: Create thumbnail previews of sheet‑specific charts for a web portal. | Extract individual chart graphics for inclusion in reports, presentations, or documentation. | Automate image generation for a data‑driven dashboard where each chart must be identifiable by its sheet name.
// AI Prompts: Generate C# code that saves all charts in an Aspose.Cells workbook as JPEG files, including both the worksheet name and chart index in each filename. | Explain how to modify the example to store chart images in a dedicated folder and handle worksheets containing multiple charts. | Provide a step‑by‑step guide for batch‑exporting charts from a large Excel file, ensuring each PNG is named with the worksheet name and a timestamp for version control.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates how to loop through all worksheets in a workbook, create a column chart on each sheet, and export each chart to a PNG image whose filename matches the worksheet name with Aspose.Cells Chart.ToImage. The workbook is then saved as an XLSX file.
class SaveChartImages
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data and a chart to each worksheet
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet ws = workbook.Worksheets[i];
            ws.Name = $"Sheet{i + 1}";

            // Populate sample data
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["B2"].PutValue(10 + i * 5);
            ws.Cells["B3"].PutValue(20 + i * 5);

            // Add a column chart
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = ws.Charts[chartIndex];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Save the chart image as PNG named after the worksheet
            string imagePath = $"{ws.Name}.png";
            chart.ToImage(imagePath, ImageType.Png);
        }

        // Save the workbook (optional)
        workbook.Save("ChartsWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
