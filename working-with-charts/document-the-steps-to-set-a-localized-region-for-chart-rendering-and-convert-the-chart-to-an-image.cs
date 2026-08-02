// Title: Localize Chart Rendering with Workbook Region and Export to PNG using Aspose.Cells for .NET
// Description: Demonstrates how to set a workbook's regional setting (e.g., Japan) in Aspose.Cells, create a column chart, apply a localized title, and export the chart as a PNG image while optionally saving the workbook.
// Keywords: Aspose.Cells | C# chart localization | Workbook region setting | CountryCode Japan | chart to image | export chart PNG | regional formatting Aspose.Cells | .NET Excel chart image
// Common Searches: Aspose.Cells set workbook locale for chart | export chart as PNG after setting region | localize chart number format Aspose.Cells | chart ToImage method .NET example | Japanese chart formatting Aspose.Cells
// Developer Intent: Apply a specific locale to a workbook so that chart axes and titles use regional formats, then generate a standalone image of the chart.
// Use Cases: Produce sales charts with Japanese number/date formats for regional reports and export them as PNG files. | Automate creation of localized charts for multiple countries and deliver both Excel files and image assets for web dashboards. | Save a workbook containing a localized chart while also providing a high‑resolution chart image for presentations or documentation.
// AI Prompts: Show how to set the workbook region to French and export a line chart as a JPEG with Aspose.Cells. | Generate code that loops through several CountryCode values, creates a chart for each region, and saves each chart as a PNG image. | Explain the impact of the Workbook.Settings.Region property on axis number formatting when converting a chart to an image.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartLocalization
{
    // Demonstrates how to set a workbook's regional setting (e.g., Japan) in Aspose.Cells, create a column chart, apply a localized title, and export the chart as a PNG image while optionally saving the workbook.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();

            // 2. Set the workbook's regional settings (e.g., Japanese locale)
            // This influences number/date formatting used during chart rendering.
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

            // 7. Optionally customize the chart (title, axis format, etc.)
            chart.Title.Text = "月別売上 (Sales by Month)";

            // 8. Convert the chart to an image file.
            // The image format is inferred from the file extension.
            string imagePath = "LocalizedChart.png";
            chart.ToImage(imagePath, ImageType.Png);

            // 9. Save the workbook (optional, to keep the chart in the Excel file)
            workbook.Save("LocalizedChartWorkbook.xlsx");

            Console.WriteLine($"Chart image saved to: {Path.GetFullPath(imagePath)}");
            Console.WriteLine("Workbook saved to: LocalizedChartWorkbook.xlsx");
        }
    }
}
