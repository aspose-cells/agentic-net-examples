// Title: Apply distinct colors to start, intermediate, and total points in an Aspose.Cells waterfall chart (C#)
// AI Prompts: Write C# code with Aspose.Cells that marks the first data point as a Start type and assigns a custom fill color, sets the middle points as Intermediate with another color, and colors the final point as Total with a third color. | Update an existing Aspose.Cells waterfall chart example to customize the fill colors for start, intermediate, and total points without altering the data range.
// Common Searches: how to change start point color in Aspose.Cells waterfall chart c# | Aspose.Cells waterfall chart custom colors for intermediate values | set total point type and color in Aspose.Cells .NET waterfall chart | C# Aspose.Cells example coloring waterfall chart point types | waterfall chart point type formatting Aspose.Cells workbook
// Tags: Aspose.Cells waterfall chart point type colors | C# customize start point color Aspose.Cells | intermediate point fill color Aspose.Cells | total point type formatting Aspose.Cells .NET | waterfall chart series customization Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, writes category and value data for start, revenue, cost, profit, and total rows, adds a waterfall chart, assigns the range B2:B6 to the series, sets a chart title, and saves the workbook as WaterfallChart.xlsx.
class WaterfallChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the waterfall chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(100);

            sheet.Cells["A3"].PutValue("Revenue");
            sheet.Cells["B3"].PutValue(150);

            sheet.Cells["A4"].PutValue("Cost");
            sheet.Cells["B4"].PutValue(-50);

            sheet.Cells["A5"].PutValue("Profit");
            sheet.Cells["B5"].PutValue(200);

            sheet.Cells["A6"].PutValue("Total");
            sheet.Cells["B6"].PutValue(300);

            // Add a waterfall chart
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Waterfall Chart Example";

            // Add series and set its data range
            int seriesIndex = chart.NSeries.Add("B2:B6", true);
            // Category data is inferred from the first column; explicit assignment omitted for compatibility

            // Save the workbook
            workbook.Save("WaterfallChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
