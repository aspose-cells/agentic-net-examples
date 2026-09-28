// Title: Insert a multi‑series line chart for quarterly product sales into an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a line chart showing quarterly sales for several products on the same worksheet, including setting the X‑axis categories and series names. | Demonstrate how to add a line chart, configure its title, assign category data, add multiple data series, and save the workbook as an .xlsx file using Aspose.Cells in C#.
// Common Searches: how to add a line chart with multiple data series in Aspose.Cells C# | Aspose.Cells create quarterly sales line graph for several products | C# Aspose.Cells set category axis range for line chart | save Excel workbook with line chart using Aspose.Cells .NET | Aspose.Cells line chart marker customization C# example
// Tags: Aspose.Cells add line chart multiple series | Aspose.Cells set chart category data range | Aspose.Cells configure line chart title | Aspose.Cells save workbook with chart xlsx | Aspose.Cells line chart marker type C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a new workbook, populates it with quarterly sales data for three products, adds a line chart that uses the quarter column as the X‑axis and each product column as a separate series, sets a chart title, optionally configures markers, and saves the file as QuarterlySalesLineChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Fill sample sales data (quarters vs products)
                sheet.Cells["A1"].PutValue("Quarter");
                sheet.Cells["B1"].PutValue("Product A");
                sheet.Cells["C1"].PutValue("Product B");
                sheet.Cells["D1"].PutValue("Product C");

                // Quarter labels
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                // Sales figures
                sheet.Cells["B2"].PutValue(12000);
                sheet.Cells["B3"].PutValue(15000);
                sheet.Cells["B4"].PutValue(13000);
                sheet.Cells["B5"].PutValue(17000);

                sheet.Cells["C2"].PutValue(10000);
                sheet.Cells["C3"].PutValue(14000);
                sheet.Cells["C4"].PutValue(11000);
                sheet.Cells["C5"].PutValue(16000);

                sheet.Cells["D2"].PutValue(9000);
                sheet.Cells["D3"].PutValue(12000);
                sheet.Cells["D4"].PutValue(11500);
                sheet.Cells["D5"].PutValue(15000);

                // Add a line chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart title
                chart.Title.Text = "Quarterly Sales Comparison";

                // Set category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A5";

                // Add series for each product
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries[0].Name = "Product A";

                chart.NSeries.Add("C2:C5", true);
                chart.NSeries[1].Name = "Product B";

                chart.NSeries.Add("D2:D5", true);
                chart.NSeries[2].Name = "Product C";

                // Optional: format chart markers (if supported by the library version)
                // Uncomment the following lines if Marker.Type is available in your Aspose.Cells version.
                // chart.NSeries[0].Marker.Type = MarkerType.Circle;
                // chart.NSeries[1].Marker.Type = MarkerType.Circle;
                // chart.NSeries[2].Marker.Type = MarkerType.Circle;

                // Save the workbook
                string outputPath = "QuarterlySalesLineChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
