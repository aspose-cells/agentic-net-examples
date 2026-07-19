// Title: Set Column Chart Data Labels to Inside End Position with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a column chart, enable data labels, and position those labels inside the column ends using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# chart data labels | InsideEnd label position | column chart Aspose | Excel chart customization | set data label position | .NET Excel library | chart label placement
// Common Searches: Aspose.Cells set column chart label InsideEnd | C# place chart data labels inside end | how to change data label position in Aspose.Cells | column chart label placement .NET | Aspose.Cells chart label visibility
// Developer Intent: Place column chart data labels at the Inside End location using Aspose.Cells.
// Use Cases: Display values inside each column for compact financial reports. | Create sales dashboards where labels do not overlap axis titles. | Generate Excel spreadsheets with cleanly positioned data labels for better readability. | Automate chart styling in bulk processing of Excel files.
// AI Prompts: Write C# code with Aspose.Cells to set data label position to InsideEnd for a column chart series. | Explain how to adjust data label positions for multiple series in an Aspose.Cells chart. | Show how to toggle data label visibility and choose different positions (InsideEnd, OutsideEnd, Center) in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a column chart, enable data labels, and position those labels inside the column ends using Aspose.Cells for C#.
class SetDataLabelPositionInsideEnd
{
    static void Main()
    {
        // Create a new workbook
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

        // Set the data range for the series
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        chart.NSeries[0].DataLabels.ShowValue = true;

        // Set data label position to InsideEnd for better readability
        chart.NSeries[0].DataLabels.Position = LabelPositionType.InsideEnd;

        // Save the workbook
        workbook.Save("ColumnChart_InsideEndDataLabels.xlsx");
    }
}
