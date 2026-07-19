// Title: Set distinct colors for column and line series in an Aspose.Cells combo chart (C#/.NET)
// Description: Creates a workbook, adds sales and target data, builds a combo chart with a column series and a line series, applies a blue fill to the columns and a red solid line to the line series, then saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# combo chart color | column series fill color | line series color | chart series formatting | Excel chart customization | .NET chart styling | ComboChart Aspose.Cells | ChartSeriesColor C#
// Common Searches: Aspose.Cells change column series color combo chart | Set line series color in Aspose.Cells chart C# | How to format combo chart series colors with Aspose.Cells | Customize colors for column and line series in Excel using Aspose
// Developer Intent: Generate a combo chart and assign a unique fill color to the column series and a distinct line color to the line series using Aspose.Cells for .NET.
// Use Cases: Financial dashboards that need separate visual cues for revenue (columns) and forecast (line). | Sales reports where actual sales columns are highlighted in blue and target trends are shown in red. | Automated Excel generation for presentations that require clear differentiation between multiple data series.
// AI Prompts: Write C# code with Aspose.Cells to create a combo chart where the column series is blue and the line series is red. | Explain how to set fill colors for column series and line colors for line series in an Aspose.Cells combo chart. | Show how to apply solid line formatting to a line series in a combo chart using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsComboChartDemo
{
    // Creates a workbook, adds sales and target data, builds a combo chart with a column series and a line series, applies a blue fill to the columns and a red solid line to the line series, then saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Category column
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Column series data
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(200);

            // Line series data
            sheet.Cells["C1"].PutValue("Target");
            sheet.Cells["C2"].PutValue(130);
            sheet.Cells["C3"].PutValue(140);
            sheet.Cells["C4"].PutValue(170);
            sheet.Cells["C5"].PutValue(210);

            // Add a combo chart (initially a Column chart)
            int chartIdx = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIdx];

            // Add the column series (first series)
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries[0].Name = "Sales";

            // Add the line series (second series) and set its chart type to Line
            chart.NSeries.Add("C2:C5", true);
            chart.NSeries[1].Name = "Target";
            chart.NSeries[1].Type = ChartType.Line; // Convert second series to line

            // Set distinct fill color for the column series (index 0)
            chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(79, 129, 189); // a blue shade

            // Set distinct line color for the line series (index 1)
            chart.NSeries[1].SeriesLines.Color = Color.Red;
            chart.NSeries[1].SeriesLines.FormattingType = ChartLineFormattingType.Solid;

            // Optional: set category axis data (already set by Add with true)
            chart.NSeries.CategoryData = "A2:A5";

            // Save the workbook
            workbook.Save("ComboChartDistinctColors.xlsx");
        }
    }
}
