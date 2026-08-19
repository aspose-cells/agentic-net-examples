// Title: Rotate All Chart Tick Labels 90° Across Worksheets with Aspose.Cells for .NET (C#)
// Description: C# example that loads or creates a workbook, iterates through every worksheet and each chart, and sets both the CategoryAxis and ValueAxis tick‑label direction to ChartTextDirectionType.Rotate90. The workbook is then saved with the rotated labels applied to all charts.
// Keywords: Aspose.Cells C# chart rotate labels | Rotate90 tick labels Aspose.Cells | set chart axis label direction .NET | iterate worksheets charts Aspose.Cells | CategoryAxis TickLabels Rotate90 | ValueAxis TickLabels Rotate90 | Excel chart label orientation example | Aspose.Cells chart axis formatting
// Common Searches: how to rotate chart axis labels 90 degrees Aspose.Cells | C# set tick label direction for all charts in workbook | Aspose.Cells iterate worksheets and modify chart axes | rotate category and value axis labels to 90° in .NET | apply Rotate90 to chart tick labels across multiple sheets
// Developer Intent: Apply a 90° rotation to the tick labels of every chart’s category and value axes in all worksheets of an Excel file using Aspose.Cells for .NET.
// Use Cases: Standardize vertical label orientation in financial dashboards where column names are long. | Prepare charts for narrow‑column printing or PDF export by rotating axis labels to save space. | Ensure consistent label rotation when generating reports that contain multiple worksheets with charts.
// AI Prompts: Generate C# code with Aspose.Cells that rotates tick labels of all charts in a workbook to 90 degrees. | Show how to loop through each worksheet and set both CategoryAxis and ValueAxis TickLabels.DirectionType to Rotate90. | Explain how to modify existing chart axis label direction without recreating the chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that loads or creates a workbook, iterates through every worksheet and each chart, and sets both the CategoryAxis and ValueAxis tick‑label direction to ChartTextDirectionType.Rotate90. The workbook is then saved with the rotated labels applied to all charts.
class SetTickLabelDirection
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Example data for demonstration
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a sample chart
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Iterate through all worksheets
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Iterate through all charts in the worksheet
            foreach (Chart ch in ws.Charts)
            {
                // Set tick label direction to Rotate90 for category axis
                ch.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Rotate90;

                // Set tick label direction to Rotate90 for value axis
                ch.ValueAxis.TickLabels.DirectionType = ChartTextDirectionType.Rotate90;
            }
        }

        // Save the workbook
        workbook.Save("TickLabelsDirectionRotate90.xlsx");
    }
}
