// Title: C# – Position Chart Legend at Bottom with Aspose.Cells for .NET
// Description: This Aspose.Cells for .NET example creates a workbook, fills it with sample data, adds a column chart, and moves the legend to the bottom of the chart area using the Legend.Position = LegendPositionType.Bottom property. The workbook is saved as an Excel file.
// Keywords: Aspose.Cells C# | chart legend bottom | LegendPositionType | Excel chart layout | .NET chart example | column chart Aspose.Cells | set legend position | Excel file generation | GitHub Aspose.Cells sample
// Common Searches: Aspose.Cells set legend to bottom C# | move chart legend bottom Aspose.Cells .NET | how to change legend position in Excel chart using Aspose.Cells | C# example for chart legend placement | Aspose.Cells chart layout options
// Developer Intent: Place the chart legend at the bottom of the chart area.
// Use Cases: Generate Excel reports where the legend appears below the chart for clearer presentation. | Create reusable chart templates with bottom‑positioned legends across different chart types. | Automate workbook creation that requires balanced chart layouts in .NET applications.
// AI Prompts: Show C# code to set a pie chart legend at the bottom using Aspose.Cells. | Provide examples of moving the legend to top, left, or right positions in an Excel chart with Aspose.Cells. | Explain each value of the LegendPositionType enum and when to use them.

using Aspose.Cells;
using Aspose.Cells.Charts;

// This Aspose.Cells for .NET example creates a workbook, fills it with sample data, adds a column chart, and moves the legend to the bottom of the chart area using the Legend.Position = LegendPositionType.Bottom property. The workbook is saved as an Excel file.
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

        // Position the legend at the bottom of the chart area for a balanced layout
        chart.Legend.Position = LegendPositionType.Bottom;

        // Save the workbook to a file
        workbook.Save("ChartWithBottomLegend.xlsx");
    }
}
