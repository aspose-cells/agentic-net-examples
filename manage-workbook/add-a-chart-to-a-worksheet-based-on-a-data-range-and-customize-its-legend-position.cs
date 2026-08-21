// Title: Add a Column Chart and Position Its Legend at the Bottom with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new workbook, fill cells A1:B5 with sample data, insert a column chart, bind the chart to the range A1:B5, move the legend to the bottom of the chart area, disable overlay, and save the file as ChartWithLegend.xlsx using Aspose.Cells in C#.
// Keywords: Aspose.Cells C# chart | add column chart Aspose.Cells | set legend position bottom | chart legend overlay false | define chart data range Aspose.Cells | Excel automation .NET | Aspose.Cells chart example | column chart legend placement | Aspose.Cells workbook chart
// Common Searches: Aspose.Cells add column chart C# | How to move chart legend to bottom Aspose.Cells | Set chart legend not overlay Aspose.Cells .NET | Define chart data range Aspose.Cells C# | Create chart with bottom legend using Aspose.Cells
// Developer Intent: Generate a column chart from a specified cell range and place its legend below the plot area without covering the chart.
// Use Cases: Produce quarterly sales reports where a column chart visualizes revenue and the legend appears beneath the chart for clean presentation. | Automate performance dashboards that add charts to multiple worksheets, ensuring each legend is positioned outside the chart to keep data visible. | Create inventory summaries that include a column chart with a bottom legend, maintaining a consistent layout across exported Excel files.
// AI Prompts: Write C# code with Aspose.Cells to add a line chart from range C1:D10, set the legend to the right side, and prevent overlay. | Show an example that inserts three different chart types into a worksheet and aligns each legend at the top using Aspose.Cells for .NET. | Provide code to load an existing workbook, change an existing chart's legend position to bottom, and disable overlay with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a new workbook, fill cells A1:B5 with sample data, insert a column chart, bind the chart to the range A1:B5, move the legend to the bottom of the chart area, disable overlay, and save the file as ChartWithLegend.xlsx using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 5; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 10);
        }

        // Add a column chart to the worksheet (topRow, leftColumn, bottomRow, rightColumn)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 1, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart (vertical series)
        chart.SetChartDataRange("A1:B5", true);

        // Customize the legend: place it at the bottom and ensure it does not overlay the chart
        chart.Legend.Position = LegendPositionType.Bottom;
        chart.Legend.IsOverLay = false;

        // Save the workbook with the chart
        workbook.Save("ChartWithLegend.xlsx", SaveFormat.Xlsx);
    }
}
