// Title: Aspose.Cells C# – Move a Chart to cell D5 using UpperLeftRow and UpperLeftColumn
// Description: Creates a workbook, adds sample data and a column chart, then repositions the chart so its upper‑left corner aligns with cell D5 by setting ChartObject.UpperLeftRow = 4 and UpperLeftColumn = 3 (zero‑based indices). The new location is printed and the file is saved as ChartMoved.xlsx.
// Keywords: Aspose.Cells chart positioning | C# move chart to cell | UpperLeftRow UpperLeftColumn | set chart location Aspose.Cells | .NET Excel chart placement | programmatic chart relocation | Excel chart coordinates
// Common Searches: how to place an Aspose.Cells chart at D5 in C# | set chart upper left corner to a specific cell Aspose.Cells | move Excel chart programmatically using Aspose.Cells .NET | chart object UpperLeftRow UpperLeftColumn example
// Developer Intent: Reposition an existing Aspose.Cells chart so its upper‑left corner starts at cell D5.
// Use Cases: Align charts with table headers for tidy report layouts. | Build dashboards where each chart occupies a predefined cell range. | Automate Excel reports that require exact chart placement for merging with other content.
// AI Prompts: Generate C# code to move an Aspose.Cells chart to cell D5 using UpperLeftRow and UpperLeftColumn. | Explain how zero‑based row and column indices correspond to Excel cell references when positioning charts. | Show how to read the current position of a chart and then change it to a target cell in Aspose.Cells .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data and a column chart, then repositions the chart so its upper‑left corner aligns with cell D5 by setting ChartObject.UpperLeftRow = 4 and UpperLeftColumn = 3 (zero‑based indices). The new location is printed and the file is saved as ChartMoved.xlsx.
class MoveChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Fruits");
        sheet.Cells["A3"].PutValue("Vegetables");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(30);

        // Add a chart (initial position does not matter)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 0, 0, 10, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data source
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Move the chart so its upper‑left corner starts at cell D5
        // D5 corresponds to column index 3 (zero‑based) and row index 4 (zero‑based)
        chart.ChartObject.UpperLeftRow = 4;      // Row 5
        chart.ChartObject.UpperLeftColumn = 3;   // Column D

        // Optional: display the new position (adding 1 to convert back to 1‑based indices)
        Console.WriteLine($"Chart now starts at row {chart.ChartObject.UpperLeftRow + 1}, column {chart.ChartObject.UpperLeftColumn + 1}");

        // Save the workbook
        workbook.Save("ChartMoved.xlsx");
    }
}
