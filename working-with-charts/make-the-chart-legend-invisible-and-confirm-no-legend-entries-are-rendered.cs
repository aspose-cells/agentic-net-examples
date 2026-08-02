// Title: Hide Chart Legend in Aspose.Cells for .NET and Verify No Legend Rendering
// Description: C# example that creates a workbook, adds sample data, inserts a column chart, disables the legend with ShowLegend = false, calculates the chart, prints the ShowLegend flag and the LegendEntries count (which remains >0 but is not rendered), and saves the result as ChartWithoutLegend.xlsx.
// Keywords: Aspose.Cells hide chart legend | ShowLegend false C# | Aspose.Cells legend visibility | verify legend entries Aspose.Cells | remove Excel chart legend programmatically | Aspose.Cells chart formatting | C# Excel chart example
// Common Searches: how to hide chart legend using Aspose.Cells .NET | Aspose.Cells ShowLegend property example | check legend entries after hiding legend Aspose.Cells | C# hide Excel chart legend programmatically | Aspose.Cells verify legend not rendered
// Developer Intent: Make the chart legend invisible and confirm that no legend entries are rendered in the output file.
// Use Cases: Generate clean reports where the legend must be hidden but still validate its data programmatically. | Automate Excel dashboards that follow branding rules requiring no visible legends, while logging legend entry counts for debugging. | Create Excel files for downstream processing where legends are unnecessary, ensuring they do not affect layout or rendering.
// AI Prompts: Write C# code with Aspose.Cells that hides a chart legend using ShowLegend = false and confirms that LegendEntries are not displayed after chart.Calculate(). | Explain how the ShowLegend property influences chart rendering in Aspose.Cells and how to programmatically verify that the legend is not rendered. | Suggest alternative methods in Aspose.Cells to suppress a chart legend, such as clearing LegendEntries or setting legend formatting to transparent, and compare them with using ShowLegend.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that creates a workbook, adds sample data, inserts a column chart, disables the legend with ShowLegend = false, calculates the chart, prints the ShowLegend flag and the LegendEntries count (which remains >0 but is not rendered), and saves the result as ChartWithoutLegend.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Hide the legend using the ShowLegend property
        chart.ShowLegend = false;

        // Verify that the legend is hidden
        Console.WriteLine("Chart.ShowLegend after setting to false: " + chart.ShowLegend);

        // Calculate the chart to ensure internal structures are built
        chart.Calculate();

        // Check the number of legend entries (they exist but are not rendered because the legend is hidden)
        Console.WriteLine("Legend entries count (should be >0 but not displayed): " + chart.Legend.LegendEntries.Count);

        // Save the workbook
        workbook.Save("ChartWithoutLegend.xlsx");
    }
}
