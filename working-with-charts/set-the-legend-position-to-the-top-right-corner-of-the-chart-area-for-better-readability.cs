// Title: Aspose.Cells C# – Set Chart Legend to Top‑Right Corner
// Description: Demonstrates how to create a workbook, add a column chart, and position its legend in the top‑right corner using the LegendPositionType.Corner enum. The example saves the result as LegendTopRight.xlsx.
// Keywords: Aspose.Cells chart legend position | C# set legend top right | LegendPositionType Corner | Aspose.Cells move legend | .NET Excel chart legend | chart legend placement Aspose
// Common Searches: Aspose.Cells set legend top right C# | how to move chart legend to corner in Aspose.Cells | C# Aspose.Cells legend position enum | place Excel chart legend at top‑right with Aspose | Aspose.Cells chart legend placement examples
// Developer Intent: Position a chart legend in the top‑right corner of the chart area.
// Use Cases: Generate Excel reports where legends are consistently placed at the top‑right for better readability. | Create automated dashboards that require a uniform legend location across multiple charts. | Apply corporate styling rules that mandate corner‑anchored legends in all exported spreadsheets.
// AI Prompts: Write C# code with Aspose.Cells to add a line chart and set its legend to the top‑right corner. | Show how to loop through every chart in a workbook and set each legend position to Corner. | Explain the effect of each LegendPositionType enum value on legend placement in Aspose.Cells charts.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a column chart, and position its legend in the top‑right corner using the LegendPositionType.Corner enum. The example saves the result as LegendTopRight.xlsx.
class SetLegendTopRight
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Insert a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B4", true);

        // Set the legend position to the top‑right corner of the chart area
        chart.Legend.Position = LegendPositionType.Corner;

        // Save the workbook
        workbook.Save("LegendTopRight.xlsx");
    }
}
