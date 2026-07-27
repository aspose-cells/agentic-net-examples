// Title: Rotate all chart axis tick labels 90° across worksheets with Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds a column chart, then iterates through every worksheet and each chart to set CategoryAxis and ValueAxis tick label DirectionType to ChartTextDirectionType.Rotate90 before saving the file.
// Keywords: Aspose.Cells | C# | .NET | chart axis label rotation | ChartTextDirectionType.Rotate90 | CategoryAxis.TickLabels | ValueAxis.TickLabels | iterate worksheets charts | multiple charts Aspose.Cells | rotate X axis labels | rotate Y axis labels
// Common Searches: C# Aspose.Cells rotate chart axis labels 90 degrees | set tick label direction for all charts in workbook Aspose.Cells | iterate worksheets and change chart label orientation .NET | ChartTextDirectionType.Rotate90 example Aspose.Cells | how to rotate X and Y axis labels for multiple charts
// Developer Intent: Programmatically set the tick label direction of both category and value axes to a 90° rotation for every chart in all worksheets of an Excel workbook.
// Use Cases: Standardize axis label orientation in multi‑sheet financial reports before distribution. | Improve readability of dense category labels in dashboards that contain charts on each worksheet. | Prepare Excel files for printing where vertical labels prevent overlap.
// AI Prompts: Generate C# code using Aspose.Cells that loops through all worksheets and sets ChartTextDirectionType.Rotate90 for CategoryAxis and ValueAxis tick labels. | Explain how to modify existing chart axis label direction without recreating the chart in an Aspose.Cells workbook. | Show how to verify that the tick label direction has been applied to each chart after saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that creates a workbook, adds a column chart, then iterates through every worksheet and each chart to set CategoryAxis and ValueAxis tick label DirectionType to ChartTextDirectionType.Rotate90 before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Iterate through all worksheets in the workbook
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Iterate through all charts in the current worksheet
            foreach (Chart ch in ws.Charts)
            {
                // Set tick label direction for the category (X) axis
                ch.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Rotate90;

                // Set tick label direction for the value (Y) axis
                ch.ValueAxis.TickLabels.DirectionType = ChartTextDirectionType.Rotate90;
            }
        }

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
