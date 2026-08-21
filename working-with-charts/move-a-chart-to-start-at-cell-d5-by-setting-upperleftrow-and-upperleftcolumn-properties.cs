// Title: C# – Move an Aspose.Cells chart to start at cell D5
// Description: Shows how to create a workbook, add sample data, insert a column chart, and reposition its upper‑left corner to cell D5 by assigning UpperLeftRow = 4 and UpperLeftColumn = 3 on the ChartObject, then save the file.
// Keywords: Aspose.Cells | C# | chart positioning | UpperLeftRow | UpperLeftColumn | move chart to cell | Excel chart placement | programmatic chart location
// Common Searches: Aspose.Cells set chart position C# | How to move chart to D5 Aspose.Cells | ChartObject UpperLeftRow UpperLeftColumn example | Reposition Excel chart with Aspose.Cells .NET | Change chart location after creation Aspose.Cells
// Developer Intent: Place the chart’s upper‑left corner at cell D5 programmatically.
// Use Cases: Generate a report where the chart must align with a specific cell for layout consistency. | Adjust chart locations dynamically based on user‑defined positions. | Batch‑process worksheets to align multiple charts with designated cells.
// AI Prompts: Write C# code using Aspose.Cells to move an existing chart so its top‑left corner starts at cell D5. | Show how to calculate zero‑based indices for Excel cells when setting UpperLeftRow and UpperLeftColumn. | Explain how to reposition a chart after creation in Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add sample data, insert a column chart, and reposition its upper‑left corner to cell D5 by assigning UpperLeftRow = 4 and UpperLeftColumn = 3 on the ChartObject, then save the file.
class MoveChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data for the chart (optional but useful for a visible chart)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a chart at an initial position
        int chartIndex = sheet.Charts.Add(ChartType.Column, 0, 0, 10, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Move the chart so its upper‑left corner starts at cell D5.
        // Row and column indices are zero‑based: D = 3, 5 = 4.
        chart.ChartObject.UpperLeftRow = 4;      // Row index for row 5
        chart.ChartObject.UpperLeftColumn = 3;   // Column index for column D

        // Save the workbook with the moved chart
        workbook.Save("ChartMoved.xlsx");
    }
}
