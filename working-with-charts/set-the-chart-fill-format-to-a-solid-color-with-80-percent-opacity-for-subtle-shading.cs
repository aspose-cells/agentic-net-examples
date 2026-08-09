// Title: Set Chart Area Solid Fill with 80% Opacity Using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, and applies a solid LightGray fill with 20% transparency (80% opacity) to the chart area, then saves as ChartWithSolidFill.xlsx.
// Keywords: Aspose.Cells chart fill | C# chart area opacity | solid fill Aspose.Cells | chart transparency .NET | FillFormat SolidFill | ChartArea background color | Aspose.Cells example | column chart styling | 80 percent opacity chart | Aspose.Cells C# tutorial
// Common Searches: how to set chart area fill opacity in Aspose.Cells C# | Aspose.Cells solid fill with transparency example | change background color of chart area Aspose.Cells .NET | set chart transparency using FillFormat in Aspose.Cells | C# code for semi‑transparent chart fill Aspose.Cells
// Developer Intent: Apply a solid LightGray fill with 80 % opacity to a chart’s area.
// Use Cases: Add a subtle gray background to charts for slide decks or reports. | Standardize semi‑transparent chart shading across multiple worksheets in a financial model. | Match corporate branding by using a consistent, partially opaque chart area color.
// AI Prompts: Generate C# code that uses Aspose.Cells to set a chart area’s FillFormat to a solid color with 80% opacity and save the workbook. | Show an Aspose.Cells example that changes a chart’s background to LightGray with 20% transparency (0.2) in .NET. | Explain the relationship between the Transparency property and opacity when styling a chart area with FillFormat in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartFillExample
{
    // Creates a workbook, adds sample data, inserts a column chart, and applies a solid LightGray fill with 20% transparency (80% opacity) to the chart area, then saves as ChartWithSolidFill.xlsx.
    public class Program
    {
        public static void Main()
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

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ----- Set chart fill format -----
            // Use the solid fill of the chart area and set a color with 80% opacity (20% transparency)
            chart.ChartArea.Area.FillFormat.FillType = FillType.Solid;               // Ensure solid fill
            chart.ChartArea.Area.FillFormat.SolidFill.Color = Color.LightGray;      // Choose a subtle shade
            chart.ChartArea.Area.FillFormat.SolidFill.Transparency = 0.2;           // 0.2 = 20% transparent (80% opaque)

            // Save the workbook to a file
            workbook.Save("ChartWithSolidFill.xlsx");
        }
    }
}
