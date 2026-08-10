// Title: Aspose.Cells .NET: Position Chart Legend Bottom‑Right and Apply Calibri 10 Font
// Description: C# example that creates a workbook, adds a column chart, shows the legend, moves it to the bottom‑right corner using LegendPositionType.Corner, and sets the legend font to Calibri size 10 before saving the file.
// Keywords: Aspose.Cells chart legend position | Aspose.Cells legend bottom right | Aspose.Cells set legend font | Calibri 10 legend Aspose.Cells | .NET Excel chart formatting | LegendPositionType.Corner
// Common Searches: Aspose.Cells move chart legend to bottom right | How to change chart legend font in Aspose.Cells .NET | LegendPositionType.Corner example | Set legend font Calibri using Aspose.Cells | Excel chart legend formatting with Aspose.Cells
// Developer Intent: Place the chart legend in the bottom‑right corner of the plot area and apply a Calibri 10‑point font.
// Use Cases: Standardize corporate Excel reports so every chart legend appears in the bottom‑right corner with the company‑specified Calibri 10 style. | Automate generation of sales dashboards where legends must be positioned consistently for readability across dozens of workbooks. | Create reusable chart templates for financial models that enforce legend placement and typography without manual editing.
// AI Prompts: Generate C# code with Aspose.Cells that moves a chart legend to the bottom‑right corner and sets the font to Calibri 10. | Show how to format an existing Aspose.Cells chart legend's position and font in a .NET workbook. | Explain step‑by‑step how to use LegendPositionType.Corner and Font properties to style a chart legend in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that creates a workbook, adds a column chart, shows the legend, moves it to the bottom‑right corner using LegendPositionType.Corner, and sets the legend font to Calibri size 10 before saving the file.
class ChartLegendExample
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

        // Ensure the legend is visible
        chart.ShowLegend = true;

        // Position the legend at the bottom‑right corner of the plot area
        chart.Legend.Position = LegendPositionType.Corner;

        // Set legend font to Calibri, size 10
        chart.Legend.Font.Name = "Calibri";
        chart.Legend.Font.Size = 10;

        // Save the workbook
        workbook.Save("ChartWithLegendBottomRight.xlsx");
    }
}
