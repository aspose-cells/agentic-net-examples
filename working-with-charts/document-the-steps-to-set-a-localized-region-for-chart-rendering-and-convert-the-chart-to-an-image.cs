// Title: Aspose.Cells .NET: Set Workbook Region for Chart Localization and Export Chart as PNG
// Description: Demonstrates how to assign a regional setting (e.g., Japanese) to a Workbook, create a column chart, bind data, and render the chart to a PNG image using Aspose.Cells for .NET. The example also shows optional workbook saving.
// Keywords: aspocells set workbook region | chart localization aspocells | export chart to png c# | aspocells chart toimage | regional settings workbook aspocells | .net chart image conversion
// Common Searches: Aspose.Cells change chart locale | How to export Aspose.Cells chart as PNG | Set workbook region for number formatting in Aspose.Cells | Localized chart image generation Aspose.Cells .NET | C# Aspose.Cells chart rendering with culture
// Developer Intent: Apply a specific culture to a workbook so that chart labels and number formats follow that locale, then save the rendered chart as an image file.
// Use Cases: Produce sales charts formatted for Japanese conventions and embed the PNG in reports. | Generate region‑specific financial charts automatically and deliver them as images for web dashboards. | Create multilingual workbook templates where each chart is exported as an image matching the target market’s locale.
// AI Prompts: Write C# code with Aspose.Cells to set the workbook region to France and export a line chart as a JPEG image. | Provide step‑by‑step instructions for localizing chart axis labels to German and saving the chart as a BMP using Aspose.Cells for .NET. | Explain how to read a user‑selected locale at runtime, apply it to Workbook.Settings.Region, and then convert the chart to an image.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartLocalization
{
    // Demonstrates how to assign a regional setting (e.g., Japanese) to a Workbook, create a column chart, bind data, and render the chart to a PNG image using Aspose.Cells for .NET. The example also shows optional workbook saving.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();

            // 2. Set the regional (localization) settings for the workbook.
            //    This influences how numbers, dates, and other culture‑specific data are formatted
            //    when the chart is rendered.
            workbook.Settings.Region = CountryCode.Japan;   // Example: Japanese locale

            // 3. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // 4. Populate sample data that will be used by the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["B4"].PutValue(1500);

            // 5. Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // 6. Define the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // 7. Convert the chart to an image file.
            //    The image format is inferred from the file extension (PNG in this case).
            chart.ToImage("LocalizedChart.png", ImageType.Png);

            // 8. Optionally, save the workbook itself
            workbook.Save("LocalizedChartWorkbook.xlsx");

            Console.WriteLine("Chart rendered with Japanese locale and saved as image.");
        }
    }
}
