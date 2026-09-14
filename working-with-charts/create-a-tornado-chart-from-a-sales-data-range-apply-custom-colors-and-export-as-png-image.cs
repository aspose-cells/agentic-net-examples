// Title: Generate a tornado‑style dual‑bar chart from sales data and export it as a PNG image with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that populates a worksheet with region and product sales figures, adds a bar chart that mimics a tornado chart, applies red and blue colors to the two series, sets a chart title, and saves the chart as a PNG image using Aspose.Cells. | Create a .NET example that builds a workbook, inserts sales data, creates a dual‑series bar chart with custom series colors, configures the chart title, and exports the chart to a PNG file while also saving the workbook.
// Common Searches: c# aspnet create tornado chart from sales data using Aspose.Cells | how to set custom colors for series in Aspose.Cells bar chart | export Aspose.Cells chart as PNG file in C# | simulate tornado chart with bar chart in Aspose.Cells | Aspose.Cells chart title set programmatically c#
// Tags: Aspose.Cells create tornado chart | Aspose.Cells bar chart custom series colors | Aspose.Cells export chart to PNG | Aspose.Cells populate worksheet sales data | Aspose.Cells set chart title

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a new workbook, fills it with region and product sales data, adds a bar chart that functions as a tornado chart, assigns red and blue colors to the two series, sets a chart title, and exports the chart as a PNG image while also saving the workbook as an XLSX file.
class TornadoChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sales data (Region, Product A, Product B)
            sheet.Cells["A1"].PutValue("Region");
            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["C1"].PutValue("Product B");

            sheet.Cells["A2"].PutValue("North");
            sheet.Cells["A3"].PutValue("South");
            sheet.Cells["A4"].PutValue("East");
            sheet.Cells["A5"].PutValue("West");

            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(90);
            sheet.Cells["B5"].PutValue(130);

            sheet.Cells["C2"].PutValue(100);
            sheet.Cells["C3"].PutValue(140);
            sheet.Cells["C4"].PutValue(80);
            sheet.Cells["C5"].PutValue(110);

            // Add a Bar chart (used to simulate a Tornado chart)
            int chartIndex = sheet.Charts.Add(ChartType.Bar, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set data ranges for the two series
            chart.NSeries.Add("B2:B5", true); // Product A
            chart.NSeries.Add("C2:C5", true); // Product B
            chart.NSeries.CategoryData = "A2:A5";

            // Apply custom colors to each series
            chart.NSeries[0].Area.ForegroundColor = Color.Red;   // Product A
            chart.NSeries[1].Area.ForegroundColor = Color.Blue;  // Product B

            // Set chart title
            chart.Title.Text = "Sales Tornado Chart";

            // Export the chart as a PNG image
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Default image format is PNG; explicit setting omitted for compatibility
                HorizontalResolution = 96,
                VerticalResolution = 96
            };

            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    chart.ToImage(ms, imgOptions);
                    File.WriteAllBytes("TornadoChart.png", ms.ToArray());
                }
            }
            catch (Exception imgEx)
            {
                Console.WriteLine("Error exporting chart image: " + imgEx.Message);
            }

            // Save the workbook (optional)
            workbook.Save("TornadoChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
