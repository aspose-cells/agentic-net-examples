// Title: Set Workbook Locale and Export Chart as PNG with Aspose.Cells for .NET
// Description: Creates a new Workbook, assigns a Japanese locale via Workbook.Settings.Region (which changes number and date formatting for charts), fills sample data, adds a column chart, defines its data range, and saves the chart as a PNG image. The workbook can also be saved for verification.
// Keywords: Aspose.Cells | .NET chart export | set workbook locale | Workbook.Settings.Region | chart to PNG | chart image conversion | regional formatting Aspose | localized chart rendering | export Aspose chart image
// Common Searches: How to set a locale for a workbook in Aspose.Cells | Export Aspose.Cells chart as PNG image | Apply Japanese regional settings to charts in .NET | Workbook.Settings.Region example code | Convert Aspose chart to image with specific culture
// Developer Intent: Apply a specific locale to a workbook so that charts render with that culture's formatting and then save the chart as an image file.
// Use Cases: Produce a sales column chart with Japanese number formatting and embed the PNG in a localized report. | Generate chart images for a multilingual dashboard by switching Workbook.Settings.Region per user language before rendering each chart. | Validate regional formatting by comparing the saved workbook (with locale applied) against the exported chart image.
// AI Prompts: Write C# code that sets the workbook region to France, creates a pie chart, and saves the chart as a JPEG using Aspose.Cells. | Explain how to change the workbook's regional setting at runtime based on a user's language preference before exporting a chart image. | Show how to export several charts from one workbook, each using a different locale, to separate PNG files.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartLocalization
{
    // Creates a new Workbook, assigns a Japanese locale via Workbook.Settings.Region (which changes number and date formatting for charts), fills sample data, adds a column chart, defines its data range, and saves the chart as a PNG image. The workbook can also be saved for verification.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();

            // 2. Set the workbook's regional settings (e.g., Japanese locale)
            // This influences number/date formatting during chart rendering.
            workbook.Settings.Region = CountryCode.Japan;

            // 3. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // 4. Populate sample data for the chart
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

            // 7. Convert the chart to an image file (PNG format)
            // The image type is inferred from the ImageType enum.
            chart.ToImage("LocalizedChart.png", ImageType.Png);

            // 8. Optionally save the workbook (not required for the image, but useful for verification)
            workbook.Save("LocalizedChartWorkbook.xlsx");

            Console.WriteLine("Chart rendered with Japanese regional settings and saved as image.");
        }
    }
}
