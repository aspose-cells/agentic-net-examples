// Title: Duplicate a Column Chart, Change Legend Position, and Save Workbook – Aspose.Cells C# Example
// Description: Demonstrates how to create a workbook with sample data, add a column chart, duplicate the chart, set a different legend position for the copy, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart example | duplicate chart Aspose.Cells | chart legend position .NET | column chart Excel | copy chart workbook | Aspose.Cells tutorial | Excel chart manipulation | Aspose.Cells legend bottom | save multiple charts
// Common Searches: Aspose.Cells duplicate chart C# | how to change chart legend position Aspose.Cells | copy chart and modify legend Aspose.Cells .NET | create two charts with same data range Aspose.Cells | Aspose.Cells example for chart duplication
// Developer Intent: Copy an existing chart, assign a new legend placement, and persist both charts in the same Excel workbook.
// Use Cases: Financial reports that need a primary chart with a right‑hand legend and a summary chart with a bottom legend. | Side‑by‑side visual comparisons where the duplicated chart uses a different legend layout for clearer presentation. | Automated Excel generation that includes both original and duplicated charts with distinct legend positions.
// AI Prompts: Generate C# code with Aspose.Cells to clone a column chart, set the duplicate's legend to Bottom, and save the workbook. | Show an Aspose.Cells example that copies a chart, moves it to a new location, changes its legend position, and writes the file to disk.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook with sample data, add a column chart, duplicate the chart, set a different legend position for the copy, and save the file using Aspose.Cells for .NET.
class DuplicateChartExample
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

        // Add the original chart
        int originalChartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart originalChart = sheet.Charts[originalChartIndex];
        originalChart.SetChartDataRange("A1:B4", true);
        // Keep the default legend position (right)
        originalChart.Legend.Position = LegendPositionType.Right;

        // Duplicate the chart by adding a new chart with the same data range
        int duplicateChartIndex = sheet.Charts.Add(ChartType.Column, 22, 0, 37, 8);
        Chart duplicateChart = sheet.Charts[duplicateChartIndex];
        duplicateChart.SetChartDataRange("A1:B4", true);
        // Modify the legend position of the duplicated chart
        duplicateChart.Legend.Position = LegendPositionType.Bottom;

        // Save the workbook containing both charts
        workbook.Save("DuplicatedChart.xlsx");
    }
}
