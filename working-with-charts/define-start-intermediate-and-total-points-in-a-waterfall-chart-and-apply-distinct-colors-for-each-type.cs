// Title: How to create a Waterfall chart with start, intermediate, and total points and assign custom colors using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom color for start point in a waterfall chart | C# code to color positive and negative intermediate points in Aspose.Cells waterfall chart | Define total point color in Aspose.Cells waterfall chart example | Create and style a waterfall chart with Aspose.Cells for .NET | Aspose.Cells waterfall chart series point formatting C#
// Tags: waterfall chart point color customization Aspose.Cells | custom start point styling Aspose.Cells C# | positive negative intermediate point colors Aspose.Cells | total point color formatting Aspose.Cells | export waterfall chart to XLSX Aspose.Cells | series point foreground color API Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace WaterfallChartDemo
{
    // The example creates a new workbook, fills it with category and value data for start, intermediate, and total points, adds a Waterfall chart, applies distinct foreground colors to each point type (light blue for start, green for positive intermediates, red for negative intermediates, dark gray for total), and saves the result as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ------------------------------------------------------------
                // 1. Populate data for the waterfall chart
                // ------------------------------------------------------------
                // Column A – Categories
                // Column B – Values (positive for increase, negative for decrease)
                // The first row is the header
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");

                // Example data:
                // Start point (initial balance)
                cells["A2"].PutValue("Start");
                cells["B2"].PutValue(5000);

                // Intermediate points (increase / decrease)
                cells["A3"].PutValue("Sales");
                cells["B3"].PutValue(3000);
                cells["A4"].PutValue("Refunds");
                cells["B4"].PutValue(-800);
                cells["A5"].PutValue("Expenses");
                cells["B5"].PutValue(-1200);

                // Total point (final balance)
                cells["A6"].PutValue("Total");
                cells["B6"].PutValue(0); // Value will be calculated by the chart

                // ------------------------------------------------------------
                // 2. Add a Waterfall chart
                // ------------------------------------------------------------
                // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
                int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 8, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (including headers)
                chart.SetChartDataRange("A1:B6", true);

                // Set chart title
                chart.Title.Text = "Company Financial Waterfall";

                // ------------------------------------------------------------
                // 3. Apply distinct colors for each point type
                // ------------------------------------------------------------
                Series series = chart.NSeries[0];

                // Start point – light blue
                series.Points[0].Area.ForegroundColor = Color.LightBlue;

                // Intermediate points – green for positive, red for negative
                series.Points[1].Area.ForegroundColor = Color.Green;   // Sales (positive)
                series.Points[2].Area.ForegroundColor = Color.Red;     // Refunds (negative)
                series.Points[3].Area.ForegroundColor = Color.Red;     // Expenses (negative)

                // Total point – dark gray
                series.Points[4].Area.ForegroundColor = Color.DarkGray;

                // ------------------------------------------------------------
                // 4. Save the workbook
                // ------------------------------------------------------------
                string outputPath = "WaterfallChartDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
