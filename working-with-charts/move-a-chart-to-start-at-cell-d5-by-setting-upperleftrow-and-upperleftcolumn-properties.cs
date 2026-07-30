// Title: Position an Aspose.Cells chart at cell D5 using UpperLeftRow and UpperLeftColumn (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, then moves the chart so its upper‑left corner aligns with cell D5 by assigning row index 4 and column index 3 to the chart's UpperLeftRow and UpperLeftColumn properties before saving.
// Keywords: Aspose.Cells chart positioning | UpperLeftRow UpperLeftColumn | C# chart placement Excel | move chart to specific cell | Aspose.Cells .NET chart location
// Common Searches: Aspose.Cells set chart start cell D5 | C# move Aspose chart to cell | UpperLeftRow UpperLeftColumn example | position chart at D5 Aspose.Cells | chart object location Excel .NET
// Developer Intent: Place a chart so its upper‑left corner starts at cell D5.
// Use Cases: Align charts with tables in financial statements for a clean layout. | Build dashboards where each chart occupies a predefined cell range. | Adjust chart placement after data refresh to keep a consistent worksheet design.
// AI Prompts: Show C# code that moves an Aspose.Cells chart to cell D5 using UpperLeftRow and UpperLeftColumn. | Provide an example that repositions and resizes a chart to fit a specific cell block starting at D5. | Explain how zero‑based row/column indices map to Excel references when setting chart.ChartObject.UpperLeftRow and UpperLeftColumn.

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, inserts a column chart, then moves the chart so its upper‑left corner aligns with cell D5 by assigning row index 4 and column index 3 to the chart's UpperLeftRow and UpperLeftColumn properties before saving.
class MoveChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a chart with an arbitrary initial position
        int chartIndex = sheet.Charts.Add(ChartType.Column, 0, 0, 10, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart's data source
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Move the chart so its upper‑left corner starts at cell D5
        // D5 corresponds to row index 4 (zero‑based) and column index 3 (zero‑based)
        chart.ChartObject.UpperLeftRow = 4;      // Row 5
        chart.ChartObject.UpperLeftColumn = 3;   // Column D

        // Save the workbook with the repositioned chart
        workbook.Save("ChartMoved.xlsx");
    }
}
